# -*- coding: utf-8 -*-
"""
朋友圈好友画像分析 Web 服务
前端上传朋友圈截图 zip → 后端异步用 qwen3.7-plus 分析 → 生成 Excel 下载。

用法:
    py -3.12 web/server.py            # 开发 (Flask dev server, localhost:8000)
    py -3.12 web/server.py --prod     # 生产 (waitress, 0.0.0.0:8000, 公网部署用)
"""
import os, sys, json, uuid, time, threading, zipfile, shutil, argparse, tempfile
from pathlib import Path
import requests

# 把项目根目录加入 sys.path, 以便 import analyze
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import analyze  # 复用 analyze.py 的 group_friends/analyze_one/run_analysis 等

from flask import Flask, request, jsonify, send_file, render_template, abort
from flasgger import Swagger
import logging
# 关掉 werkzeug 默认的请求行(避免和 JSON 日志重复), 保留启动告警
logging.getLogger("werkzeug").setLevel(logging.WARNING)

# ========== 配置 ==========
HERE = Path(__file__).resolve().parent
TASK_ROOT = HERE / "tasks"
TEMPLATE_DIR = HERE / "templates"

MAX_CONCURRENT_TASKS = 2     # 同时最多几个分析任务在跑, 其余排队
MAX_UPLOAD_BYTES = 2 * 1024 * 1024 * 1024   # 单次上传上限 2GB
TASK_TTL_HOURS = 24          # 完成后保留时长, 超时清理
EXCEL_NAME = "好友画像分析.xlsx"

# /analyze_one 返回字段: 中文(模型输出) -> 英文(API 对外)
ANALYZE_ONE_FIELD_MAP = {
    "性别": "gender",
    "年龄段": "age_range",
    "职业或行业": "occupation",
    "兴趣标签": "interests",
    "生活状态": "life_status",
    "活跃度": "activity_level",
    "朋友圈内容摘要": "moments_summary",
    "潜在业务价值": "business_value",
    "综合画像标签": "tags",
}

# OSS URL 域名白名单(逗号分隔, 可经环境变量 OSS_DOMAIN_WHITELIST 覆盖)
_default_oss = "https://avatar-video.oss-cn-shenzhen.aliyuncs.com/"
OSS_DOMAIN_WHITELIST = [d.strip() for d in os.environ.get("OSS_DOMAIN_WHITELIST", _default_oss).split(",") if d.strip()]
MAX_IMAGE_BYTES = 20 * 1024 * 1024   # 单张截图上限 20MB

# API key 优先环境变量, 缺省回退 analyze.py 硬编码(开发期)
if os.environ.get("ANALYZE_API_KEY"):
    analyze.API_KEY = os.environ["ANALYZE_API_KEY"]

TASK_ROOT.mkdir(parents=True, exist_ok=True)

app = Flask(__name__, template_folder=str(TEMPLATE_DIR))

# Swagger UI: /apidocs  (访问 /docs 自动跳转)
Swagger(app, template={
    "info": {
        "title": "朋友圈好友画像分析 API",
        "version": "1.0",
        "description": "上传朋友圈截图 zip 批量生成好友画像 Excel, 或单好友 /analyze_one 直接返回画像 JSON。",
    },
    "consumes": ["multipart/form-data"],
    "produces": ["application/json"],
})

@app.route("/docs")
def _docs_redirect():
    from flask import redirect
    return redirect("/apidocs")

# ========== 日志 ==========
_LOG_LOCK = threading.Lock()
NOISE_PATHS = ("/apidocs", "/docs", "/flasgger_static", "/apispec_1.json", "/apispec.json")
_REQ_CTX = threading.local()

def log_event(event, **fields):
    """统一 stdout JSON 日志, 一行一条, 方便 grep/重定向。"""
    line = json.dumps({
        "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
        "event": event,
        **fields,
    }, ensure_ascii=False)
    with _LOG_LOCK:
        print(line, flush=True)

@app.before_request
def _hook_before():
    _REQ_CTX.start = time.time()
    _REQ_CTX.ip = request.remote_addr or ""

@app.after_request
def _hook_after(resp):
    try:
        path = request.path
        if any(path == p or path.startswith(p) for p in NOISE_PATHS):
            return resp
        dur = time.time() - getattr(_REQ_CTX, "start", time.time())
        summary = _summarize_request(path)
        log_event("request",
                  method=request.method, path=path, status=resp.status_code,
                  dur_s=round(dur, 2), ip=getattr(_REQ_CTX, "ip", ""),
                  **summary)
    except Exception as e:
        log_event("log_error", error=str(e))
    return resp

def _summarize_request(path):
    """各路由入参摘要(只取关键, 不打全量避免日志爆炸)。"""
    try:
        if path == "/analyze_one":
            if request.is_json:
                d = request.get_json(silent=True) or {}
                intro = (d.get("intro") or "")
                urls = d.get("oss_urls") or []
            else:
                intro = request.form.get("intro", "")
                urls = request.form.getlist("oss_urls")
            return {"intro_preview": intro[:60], "intro_len": len(intro),
                    "oss_urls_n": len(urls)}
        if path == "/update_portrait":
            d = request.get_json(silent=True) or {}
            portrait = d.get("portrait") or {}
            msgs = d.get("messages")
            return {"portrait_keys": list(portrait.keys()) if isinstance(portrait, dict) else None,
                    "messages_n": len(msgs) if isinstance(msgs, list)
                                  else ("str" if isinstance(msgs, str) else None)}
        if path == "/upload":
            f = request.files.get("file") if "file" in request.files else None
            return {"filename": (f.filename if f else None),
                    "size": request.content_length}
        if path.startswith(("/progress/", "/download/", "/cancel/")):
            return {"task_id": path.rsplit("/", 1)[-1]}
    except Exception:
        pass
    return {}

# ========== 任务状态 ==========
_TASKS_LOCK = threading.Lock()
TASKS = {}            # id -> 状态dict (内存)
_SEM = threading.Semaphore(MAX_CONCURRENT_TASKS)   # 并发上限


def new_task(shot_dir, total):
    t = {
        "id": Path(shot_dir).parent.name,        # tasks/<id>/shots -> <id>
        "status": "pending",                      # pending/running/done/failed
        "total": total,
        "done": 0,
        "fail": 0,
        "skipped": 0,
        "current": "",
        "elapsed": 0,
        "message": "",
        "log": [],                                # 最近若干条日志(滚动)
        "created_at": time.time(),
        "finished_at": None,
        "xlsx": None,                             # 完成后的 xlsx 路径
        "cancel": False,
    }
    return t


def save_status(t):
    """持久化任务状态到磁盘, 服务重启可恢复。"""
    d = Path(TASK_ROOT) / t["id"]
    st = d / "status.json"
    try:
        st.write_text(json.dumps({k: v for k, v in t.items()
                                  if k not in ("cancel",)}, ensure_ascii=False, indent=2),
                      encoding="utf-8")
    except Exception:
        pass


def load_status_from_disk():
    """启动时扫描 TASK_ROOT 恢复任务状态(仅可读字段, 不重启线程)。"""
    if not TASK_ROOT.exists():
        return
    for d in TASK_ROOT.iterdir():
        st = d / "status.json"
        if not st.exists():
            continue
        try:
            t = json.loads(st.read_text(encoding="utf-8"))
            t.setdefault("id", d.name)
            t.setdefault("cancel", False)
            # 重启后: 已完成的保持done, running/pending 的标记为 failed(中断)
            if t.get("status") in ("running", "pending"):
                t["status"] = "failed"
                t["message"] = "服务重启, 任务中断, 请重新上传"
                t["finished_at"] = time.time()
            xlsx = d / EXCEL_NAME
            if xlsx.exists():
                t["xlsx"] = str(xlsx)
            with _TASKS_LOCK:
                TASKS[t["id"]] = t
        except Exception:
            pass


def public_view(t):
    """对外暴露的任务状态(去掉内部字段)。"""
    return {
        "id": t["id"],
        "status": t["status"],
        "total": t.get("total", 0),
        "done": t.get("done", 0),
        "fail": t.get("fail", 0),
        "skipped": t.get("skipped", 0),
        "current": t.get("current", ""),
        "elapsed": t.get("elapsed", 0),
        "message": t.get("message", ""),
        "log": t.get("log", [])[-20:],
        "finished": t.get("status") in ("done", "failed"),
        "has_xlsx": bool(t.get("xlsx") and Path(t["xlsx"]).exists()),
    }


def append_log(t, line):
    # 自动加时间戳前缀 [MM-DD HH:MM:SS]
    line = f"[{time.strftime('%m-%d %H:%M:%S')}] {line}"
    t["log"].append(line)
    if len(t["log"]) > 50:
        t["log"] = t["log"][-50:]


# ========== 后台分析线程 ==========
def run_task(task_id, shot_dir, out_xlsx, out_json):
    t = TASKS.get(task_id)
    if not t:
        return
    t["status"] = "running"
    t["message"] = "排队等待中..." if not _SEM._value else "开始分析"
    save_status(t)

    # on_progress 回调: 更新内存状态 + 持久化
    def on_progress(event, data):
        t["elapsed"] = data.get("elapsed", t["elapsed"])
        if event == "start":
            t["total"] = data.get("total", 0)
            t["skipped"] = data.get("have_done", 0)
            append_log(t, f"开始: 共 {t['total']} 人待分析")
        elif event == "ok":
            t["done"] = data.get("done", t["done"])
            t["fail"] = data.get("fail", t["fail"])
            t["current"] = f"{data.get('name','')} | {data.get('job','')}"
            append_log(t, f"[OK {t['done']}/{t['total']}] {data.get('name','')}  职业={data.get('job','')}  ({data.get('dur',0)}s)")
        elif event == "fail":
            t["fail"] = data.get("fail", t["fail"])
            append_log(t, f"[失败] {data.get('name','')}: {data.get('error','')}")
        elif event == "done":
            t["done"] = data.get("done", t["done"])
            t["fail"] = data.get("fail", t["fail"])
        save_status(t)

    def cancel_check():
        return t.get("cancel", False)

    # 排队获取并发槽
    _SEM.acquire()
    try:
        t["message"] = "分析中..."
        save_status(t)
        done_c, fail_c = analyze.run_analysis(
            shot_dir, out_xlsx, out_json,
            on_progress=on_progress, cancel_check=cancel_check)
        if t.get("cancel"):
            t["status"] = "failed"
            t["message"] = "已取消"
        else:
            t["status"] = "done"
            t["message"] = f"完成: 成功 {done_c}, 失败 {fail_c}"
        t["xlsx"] = str(out_xlsx) if Path(out_xlsx).exists() else None
    except Exception as e:
        t["status"] = "failed"
        t["message"] = f"出错: {e}"
        append_log(t, f"异常: {e}")
    finally:
        t["finished_at"] = time.time()
        save_status(t)
        _SEM.release()


# ========== 路由 ==========
@app.route("/")
def index():
    """前端页面
    ---
    tags: [页面]
    summary: 上传与分析前端页面
    description: 返回上传 zip、轮询进度、下载 Excel 的单页应用。
    responses:
      200:
        description: HTML 页面
        schema: {type: string}
    """
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    """批量上传分析
    ---
    tags: [批量分析]
    summary: 上传朋友圈截图 zip, 后台异步分析
    description: |
      上传 zip(文件命名 `<微信昵称>-N.png`), 秒回任务 ID; 后台 8 路并发分析, 最多 2 个任务同时运行, 其余排队。
      用 GET /progress/{task_id} 轮询, GET /download/{task_id} 下载 Excel。
    consumes: [multipart/form-data]
    parameters:
      - name: file
        in: formData
        type: file
        description: zip 压缩包(单文件上限 2GB)
        required: true
    responses:
      200:
        description: 任务已接受
        schema:
          type: object
          properties:
            task_id: {type: string, example: "a1b2c3d4e5f6"}
            total: {type: integer, example: 142}
      400:
        description: 无文件 / 非 zip / 命名不符合 / 未识别好友
        schema: {type: object, properties: {error: {type: string}}}
      413:
        description: 超过 2GB 上限
        schema: {type: object, properties: {error: {type: string}}}
      500:
        description: 保存/解压失败
        schema: {type: object, properties: {error: {type: string}}}
    """
    if "file" not in request.files:
        return jsonify({"error": "未收到文件"}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "文件名为空"}), 400
    # 长度检查
    if request.content_length and request.content_length > MAX_UPLOAD_BYTES:
        return jsonify({"error": f"文件超过 {MAX_UPLOAD_BYTES//1024//1024}MB 上限"}), 413

    task_id = uuid.uuid4().hex[:12]
    task_dir = TASK_ROOT / task_id
    shots_dir = task_dir / "shots"
    shots_dir.mkdir(parents=True, exist_ok=True)

    zip_path = task_dir / "upload.zip"
    try:
        f.save(zip_path)
    except Exception as e:
        shutil.rmtree(task_dir, ignore_errors=True)
        return jsonify({"error": f"保存失败: {e}"}), 500

    # 解压
    try:
        with zipfile.ZipFile(zip_path) as zf:
            zf.extractall(shots_dir)
    except zipfile.BadZipFile:
        shutil.rmtree(task_dir, ignore_errors=True)
        return jsonify({"error": "不是有效的 zip 文件"}), 400
    except Exception as e:
        shutil.rmtree(task_dir, ignore_errors=True)
        return jsonify({"error": f"解压失败: {e}"}), 500
    finally:
        try:
            zip_path.unlink()
        except Exception:
            pass

    # zip 内可能有子目录, 找到含最多 -N.png 的目录作为 shots
    real_shots = find_shots_root(shots_dir)
    if real_shots is None:
        shutil.rmtree(task_dir, ignore_errors=True)
        return jsonify({"error": "zip 内未找到符合 <昵称>-N.png 命名的截图"}), 400

    # 扫描好友数
    try:
        friends = analyze.group_friends(real_shots)
    except Exception as e:
        shutil.rmtree(task_dir, ignore_errors=True)
        return jsonify({"error": f"扫描失败: {e}"}), 500

    total = len(friends)
    if total == 0:
        shutil.rmtree(task_dir, ignore_errors=True)
        return jsonify({"error": "未识别到任何好友(需 <昵称>-N.png 命名)"}), 400

    out_xlsx = task_dir / EXCEL_NAME
    out_json = task_dir / ".portrait_progress.json"

    t = new_task(str(real_shots), total)
    t["status"] = "pending"
    with _TASKS_LOCK:
        TASKS[task_id] = t
    save_status(t)

    # 起后台线程
    th = threading.Thread(target=run_task, args=(task_id, str(real_shots), str(out_xlsx), str(out_json)), daemon=True)
    th.start()

    log_event("upload_accepted", task_id=task_id, total=total, filename=f.filename)
    return jsonify({"task_id": task_id, "total": total})


@app.route("/progress/<task_id>")
def progress(task_id):
    """查询任务进度
    ---
    tags: [批量分析]
    summary: 轮询任务状态
    description: 返回任务当前进度、成功/失败计数、最近日志等; status 为 done/failed 时表示终态。
    parameters:
      - name: task_id
        in: path
        type: string
        required: true
        description: /upload 返回的 task_id
    responses:
      200:
        description: 任务状态
        schema:
          type: object
          properties:
            id: {type: string}
            status: {type: string, enum: [pending, running, done, failed]}
            total: {type: integer}
            done: {type: integer}
            fail: {type: integer}
            skipped: {type: integer}
            current: {type: string, description: 当前正在分析的好友 + 职业}
            elapsed: {type: number, description: 已耗时秒}
            message: {type: string}
            log: {type: array, items: {type: string}}
            finished: {type: boolean}
            has_xlsx: {type: boolean}
      404:
        description: 任务不存在或已过期
        schema: {type: object, properties: {error: {type: string}}}
    """
    with _TASKS_LOCK:
        t = TASKS.get(task_id)
    if not t:
        return jsonify({"error": "任务不存在或已过期"}), 404
    return jsonify(public_view(t))


@app.route("/download/<task_id>")
def download(task_id):
    """下载结果 Excel
    ---
    tags: [批量分析]
    summary: 下载生成的 Excel
    description: 任务 status=done 且 Excel 已生成时可下载。
    parameters:
      - name: task_id
        in: path
        type: string
        required: true
    responses:
      200:
        description: Excel 文件
        schema: {type: file}
      404:
        description: 任务不存在或 Excel 未就绪
        schema: {type: object, properties: {error: {type: string}}}
    """
    with _TASKS_LOCK:
        t = TASKS.get(task_id)
    if not t:
        abort(404)
    xlsx = t.get("xlsx")
    if not xlsx or not Path(xlsx).exists():
        return jsonify({"error": "Excel 尚未生成或任务未完成"}), 404
    log_event("download", task_id=task_id)
    return send_file(xlsx, as_attachment=True, download_name=EXCEL_NAME)


@app.route("/cancel/<task_id>", methods=["POST"])
def cancel(task_id):
    """取消任务
    ---
    tags: [批量分析]
    summary: 请求取消正在进行的任务
    description: 设置取消标志, 等当前好友处理完后停止后续。已完成的任务无法取消。
    parameters:
      - name: task_id
        in: path
        type: string
        required: true
    responses:
      200:
        description: 已请求取消
        schema: {type: object, properties: {ok: {type: boolean}, message: {type: string}}}
      404:
        description: 任务不存在
        schema: {type: object, properties: {error: {type: string}}}
    """
    with _TASKS_LOCK:
        t = TASKS.get(task_id)
    if not t:
        return jsonify({"error": "任务不存在"}), 404
    t["cancel"] = True
    log_event("cancel", task_id=task_id)
    return jsonify({"ok": True, "message": "已请求取消, 正在等待当前好友处理完"})


@app.route("/analyze_one", methods=["POST"])
def analyze_one_route():
    """单好友画像分析
    ---
    tags: [单好友分析]
    summary: 输入文字介绍和/或 OSS 截图 URL, 返回单好友画像 JSON
    description: |
      非批量接口。两种入参方式:
      - JSON body: `{"intro": "...", "oss_urls": ["url1", "url2"]}` (Content-Type: application/json)
      - multipart form: intro 字段 + 重复的 oss_urls 字段

      截图 URL 必须命中 OSS 域名白名单(默认 `avatar-video.oss-cn-shenzhen.aliyuncs.com`,
      可经环境变量 OSS_DOMAIN_WHITELIST 逗号分隔覆盖), 单张上限 20MB。
      **intro 与 oss_urls 不能同时为空**, 否则返回 400。
      图片临时落盘、请求结束后清理。耗时约 20-40 秒。
    consumes: [application/json]
    parameters:
      - name: body
        in: body
        required: true
        schema:
          type: object
          properties:
            intro: {type: string, description: "用户对该客户的文字介绍(可选)"}
            oss_urls:
              type: array
              items: {type: string, format: uri}
              description: "截图 OSS URL 列表(可选; 朋友圈/对话截图均可)"
    responses:
      200:
        description: 画像分析结果(英文字段)
        schema:
          type: object
          properties:
            gender: {type: string, example: "女", enum: [男, 女, 不确定]}
            age_range: {type: string, example: "25-35"}
            occupation: {type: string, example: "国际物流/货代"}
            interests: {type: array, items: {type: string}, example: ["旅游", "社交", "摄影"]}
            life_status: {type: string, example: "创业者"}
            activity_level: {type: string, example: "中,近期更新较多"}
            moments_summary: {type: string, example: "主要发布业务广告及旅游聚会照片"}
            business_value: {type: string, example: "创业者有企业财产险/重疾/理财需求"}
            tags: {type: array, items: {type: string}, example: ["事业心强", "精致生活", "商务型"]}
      400:
        description: 入参不合法(intro 和 oss_urls 同时为空 / 域名不合法 / 下载失败 / 超过 20MB)
        schema: {type: object, properties: {error: {type: string, example: "intro 和 oss_urls 不能同时为空"}}}
      500:
        description: 模型内容审核拒绝或分析异常
        schema: {type: object, properties: {error: {type: string}}}
    """
    # 入参兼容 JSON body 与 multipart form
    if request.is_json:
        data = request.get_json(silent=True) or {}
        intro = (data.get("intro") or "").strip()
        oss_urls = data.get("oss_urls") or []
        if not isinstance(oss_urls, list):
            return jsonify({"error": "oss_urls 必须是字符串数组"}), 400
    else:
        intro = request.form.get("intro", "").strip()
        oss_urls = request.form.getlist("oss_urls")
    oss_urls = [u.strip() for u in oss_urls if isinstance(u, str) and u.strip()]

    if not intro and not oss_urls:
        return jsonify({"error": "intro 和 oss_urls 不能同时为空"}), 400

    tmp_dir = None
    image_paths = []
    try:
        if oss_urls:
            tmp_dir = Path(tempfile.mkdtemp(prefix="analyze_one_"))
            for i, url in enumerate(oss_urls):
                if not any(url.startswith(p) for p in OSS_DOMAIN_WHITELIST):
                    return jsonify({"error": f"oss_url 域名不合法: {url}"}), 400
                try:
                    resp = requests.get(url, timeout=(10, 120), stream=True)
                    resp.raise_for_status()
                    content = resp.content
                except Exception as e:
                    return jsonify({"error": f"下载失败 {url}: {e}"}), 400
                if len(content) > MAX_IMAGE_BYTES:
                    return jsonify({"error": f"图片超过 {MAX_IMAGE_BYTES//1024//1024}MB 上限: {url}"}), 400
                p = tmp_dir / f"img_{i:03d}.jpg"
                p.write_bytes(content)
                image_paths.append(str(p))

        try:
            result = analyze.analyze_single(intro, image_paths)
        except RuntimeError as e:
            log_event("analyze_one_fail", intro_len=len(intro),
                      oss_urls_n=len(image_paths), error=str(e)[:200])
            return jsonify({"error": str(e)[:300]}), 500
        except Exception as e:
            log_event("analyze_one_fail", intro_len=len(intro),
                      oss_urls_n=len(image_paths), error=str(e)[:200])
            return jsonify({"error": f"分析失败: {e}"}), 500
        # 模型输出的中文 key 统一转英文, 仅保留映射内的字段
        out = {en: result.get(cn) for cn, en in ANALYZE_ONE_FIELD_MAP.items()}
        log_event("analyze_one_done", intro_len=len(intro),
                  oss_urls_n=len(image_paths),
                  occupation=out.get("occupation"), gender=out.get("gender"),
                  age_range=out.get("age_range"),
                  tags_n=len(out.get("tags") or []))
        return jsonify(out)
    finally:
        if tmp_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)


@app.route("/update_portrait", methods=["POST"])
def update_portrait_route():
    """根据对话更新客户画像
    ---
    tags: [单好友分析]
    summary: 输入现有画像 + 对话记录, 输出更新后的画像
    description: |
      用户与大模型的对话中可能提到客户新情况(新职业/新兴趣/生活状态变化/家庭变化等),
      本接口据此更新现有画像。对话未提及的字段保持原值。
      **portrait 与 messages 均必填**, 任一为空返回 400。
    consumes: [application/json]
    parameters:
      - name: body
        in: body
        required: true
        schema:
          type: object
          required: [portrait, messages]
          properties:
            portrait:
              type: object
              description: 现有画像(/analyze_one 的输出, 英文字段)
              properties:
                gender: {type: string}
                age_range: {type: string}
                occupation: {type: string}
                interests: {type: array, items: {type: string}}
                life_status: {type: string}
                activity_level: {type: string}
                moments_summary: {type: string}
                business_value: {type: string}
                tags: {type: array, items: {type: string}}
            messages:
              type: array
              description: 对话记录, 按时间顺序
              items:
                type: object
                properties:
                  role: {type: string, enum: [user, assistant], example: "user"}
                  content: {type: string, example: "Cici最近生了宝宝"}
    responses:
      200:
        description: 更新后的画像(英文字段, 同 portrait 结构)
        schema:
          type: object
          properties:
            gender: {type: string}
            age_range: {type: string}
            occupation: {type: string}
            interests: {type: array, items: {type: string}}
            life_status: {type: string}
            activity_level: {type: string}
            moments_summary: {type: string}
            business_value: {type: string}
            tags: {type: array, items: {type: string}}
      400:
        description: portrait 为空 / messages 为空 / 类型不对
        schema: {type: object, properties: {error: {type: string}}}
      500:
        description: 模型内容审核拒绝或分析异常
        schema: {type: object, properties: {error: {type: string}}}
    """
    data = request.get_json(silent=True) or {}
    portrait = data.get("portrait")
    messages = data.get("messages")

    if not isinstance(portrait, dict) or not portrait:
        return jsonify({"error": "portrait 必须是非空对象"}), 400
    if isinstance(messages, str):
        if not messages.strip():
            return jsonify({"error": "messages 不能为空"}), 400
    elif isinstance(messages, list):
        if not any(isinstance(m, dict) and str(m.get("content", "")).strip() for m in messages):
            return jsonify({"error": "messages 不能为空"}), 400
    else:
        return jsonify({"error": "messages 必须是数组或字符串"}), 400

    try:
        result = analyze.update_portrait(portrait, messages)
    except RuntimeError as e:
        log_event("update_portrait_fail",
                  portrait_keys=list(portrait.keys()),
                  messages_n=(len(messages) if isinstance(messages, list) else "str"),
                  error=str(e)[:200])
        return jsonify({"error": str(e)[:300]}), 500
    except Exception as e:
        log_event("update_portrait_fail",
                  portrait_keys=list(portrait.keys()),
                  messages_n=(len(messages) if isinstance(messages, list) else "str"),
                  error=str(e)[:200])
        return jsonify({"error": f"更新失败: {e}"}), 500
    # 算哪些字段变了 (input portrait vs model result)
    try:
        changed_keys = [k for k in set(portrait) | set(result)
                        if portrait.get(k) != result.get(k)]
    except Exception:
        changed_keys = []
    log_event("update_portrait_done",
              portrait_keys=list(portrait.keys()),
              messages_n=(len(messages) if isinstance(messages, list) else "str"),
              changed_keys=changed_keys)
    return jsonify(result)


# ========== 辅助 ==========
def find_shots_root(shots_dir):
    """zip 解压后可能含子目录, 找到 -N.png 最多的目录作为截图根。返回 Path 或 None。"""
    import re
    pat = re.compile(r".*-\d+\.png$", re.IGNORECASE)
    best_dir, best_n = None, 0
    for d, _, files in os.walk(shots_dir):
        n = sum(1 for fn in files if pat.match(fn) and not fn.startswith("."))
        if n > best_n:
            best_n, best_dir = n, Path(d)
    if best_n == 0:
        return None
    return best_dir


def cleanup_old_tasks():
    """清理超过 TTL 的已完成任务目录。"""
    now = time.time()
    if not TASK_ROOT.exists():
        return
    for d in TASK_ROOT.iterdir():
        st = d / "status.json"
        if not st.exists():
            # 无状态文件的孤儿目录, 超 1 天删
            if now - d.stat().st_mtime > TASK_TTL_HOURS * 3600:
                shutil.rmtree(d, ignore_errors=True)
            continue
        try:
            t = json.loads(st.read_text(encoding="utf-8"))
        except Exception:
            continue
        if t.get("status") in ("done", "failed") and t.get("finished_at"):
            if now - t["finished_at"] > TASK_TTL_HOURS * 3600:
                shutil.rmtree(d, ignore_errors=True)


def bg_cleaner():
    """每小时清理一次过期任务。"""
    while True:
        time.sleep(3600)
        try:
            cleanup_old_tasks()
        except Exception:
            pass


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--prod", action="store_true", help="用 waitress 生产模式, 0.0.0.0:8000")
    ap.add_argument("--port", type=int, default=8000)
    args = ap.parse_args()

    load_status_from_disk()
    threading.Thread(target=bg_cleaner, daemon=True).start()

    if args.prod:
        from waitress import serve
        print(f"生产模式启动: http://0.0.0.0:{args.port}")
        # waitress 默认 max_request_body_size=1GB, 超过返回413; 按配置上限放开
        serve(app, host="0.0.0.0", port=args.port, threads=8,
              max_request_body_size=MAX_UPLOAD_BYTES)
    else:
        print(f"开发模式启动: http://127.0.0.1:{args.port}")
        app.run(host="127.0.0.1", port=args.port, debug=False, threaded=True)
