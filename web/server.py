# -*- coding: utf-8 -*-
"""
朋友圈好友画像分析 Web 服务
前端上传朋友圈截图 zip → 后端异步用 qwen3.7-plus 分析 → 生成 Excel 下载。

用法:
    py -3.12 web/server.py            # 开发 (Flask dev server, localhost:8000)
    py -3.12 web/server.py --prod     # 生产 (waitress, 0.0.0.0:8000, 公网部署用)
"""
import os, sys, json, uuid, time, threading, zipfile, shutil, argparse
from pathlib import Path

# 把项目根目录加入 sys.path, 以便 import analyze
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import analyze  # 复用 analyze.py 的 group_friends/analyze_one/run_analysis 等

from flask import Flask, request, jsonify, send_file, render_template, abort

# ========== 配置 ==========
HERE = Path(__file__).resolve().parent
TASK_ROOT = HERE / "tasks"
TEMPLATE_DIR = HERE / "templates"

MAX_CONCURRENT_TASKS = 2     # 同时最多几个分析任务在跑, 其余排队
MAX_UPLOAD_BYTES = 2 * 1024 * 1024 * 1024   # 单次上传上限 2GB
TASK_TTL_HOURS = 24          # 完成后保留时长, 超时清理
EXCEL_NAME = "好友画像分析.xlsx"

# API key 优先环境变量, 缺省回退 analyze.py 硬编码(开发期)
if os.environ.get("ANALYZE_API_KEY"):
    analyze.API_KEY = os.environ["ANALYZE_API_KEY"]

TASK_ROOT.mkdir(parents=True, exist_ok=True)

app = Flask(__name__, template_folder=str(TEMPLATE_DIR))

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
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
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

    return jsonify({"task_id": task_id, "total": total})


@app.route("/progress/<task_id>")
def progress(task_id):
    with _TASKS_LOCK:
        t = TASKS.get(task_id)
    if not t:
        return jsonify({"error": "任务不存在或已过期"}), 404
    return jsonify(public_view(t))


@app.route("/download/<task_id>")
def download(task_id):
    with _TASKS_LOCK:
        t = TASKS.get(task_id)
    if not t:
        abort(404)
    xlsx = t.get("xlsx")
    if not xlsx or not Path(xlsx).exists():
        return jsonify({"error": "Excel 尚未生成或任务未完成"}), 404
    return send_file(xlsx, as_attachment=True, download_name=EXCEL_NAME)


@app.route("/cancel/<task_id>", methods=["POST"])
def cancel(task_id):
    with _TASKS_LOCK:
        t = TASKS.get(task_id)
    if not t:
        return jsonify({"error": "任务不存在"}), 404
    t["cancel"] = True
    return jsonify({"ok": True, "message": "已请求取消, 正在等待当前好友处理完"})


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
        serve(app, host="0.0.0.0", port=args.port, threads=8)
    else:
        print(f"开发模式启动: http://127.0.0.1:{args.port}")
        app.run(host="127.0.0.1", port=args.port, debug=False, threaded=True)
