# 朋友圈好友画像分析

用 qwen3.7-plus 多模态模型读取朋友圈截图，批量生成好友用户画像 Excel（职业 / 兴趣 / 生活状态 / 活跃度 / 潜在业务价值等标签）。

## 两种使用方式

### 1. 命令行

```bash
pip install pillow openpyxl requests
py -3.12 analyze.py --dir "<截图目录>"            # 跑全部, 可中断续跑
py -3.12 analyze.py --dir "<截图目录>" --limit 5  # 先跑5人验证
py -3.12 analyze.py --dir "<截图目录>" --only-xlsx # 仅用已有结果重新生成Excel
```

- 截图目录内文件命名 `<微信昵称>-N.png`（每人 3-4 张，N 为 1..3 或 0..3）
- 文件名去掉 `-N` 即为「微信昵称」列
- 自动排除系统账号：异常 / 微信团队 / 文件传输助手
- 结果 Excel 与临时续跑文件 `.portrait_progress.json` 生成在截图目录下，跑完保留可续跑

### 2. Web 服务（公网部署）

```bash
pip install -r web/requirements.txt
py -3.12 web/server.py             # 开发: http://127.0.0.1:8000
py -3.12 web/server.py --prod      # 生产(waitress): 0.0.0.0:8000
```

前端上传朋友圈截图 zip → 后端异步分析 → 进度页轮询 → 完成下载 Excel。

- 异步任务：上传秒回任务 ID，后台 8 路并发分析，前端轮询进度
- 多人并发上限 2（其余排队），任务状态持久化，服务重启可恢复
- 任务 ID 挂 URL（`?task=xxx`），可离开后回来查进度 / 下载
- 24h 后自动清理已完成任务

## Excel 字段

`微信昵称 | 性别 | 年龄段 | 职业或行业 | 兴趣标签 | 生活状态 | 活跃度 | 朋友圈内容摘要 | 潜在业务价值 | 综合画像标签`

## 备注

- 单批 700+ 人约 1-2 小时
- 偶发截图触发 API 内容审核拒绝的好友会在 Excel 标记失败原因
