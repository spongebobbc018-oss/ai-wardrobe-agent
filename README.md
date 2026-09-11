# 衣橱顾问 Agent

可独立恢复的 AI 穿搭导购项目。当前版本包含：

- 商品 Excel 自动导入与库存、上架、尺码、预算、禁忌品类硬校验；
- 对话式搭配推荐、连续调整、场景冲突与无匹配状态；
- 匿名设备档案、我的衣橱与穿衣日历；
- 本地 SQLite 开发存储、可审计 SQL 迁移与 Docker 部署入口；
- 不依赖秒哒运行。

## 本地启动

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install -r agent-backend/requirements.txt
uvicorn app.main:app --app-dir agent-backend --host 127.0.0.1 --port 8080
```

打开 `http://127.0.0.1:8080`。首次启动会从 `AI穿搭导购_商品更新源.xlsx` 导入 70 条商品。

## Docker 启动

```bash
docker compose up --build
```

打开 `http://127.0.0.1:8080`。容器运行数据保存在 `wardrobe_data` 卷中，删除容器不会删除该卷。

## 生产环境变量

复制 `agent-backend/.env.example` 并按部署环境设置：

- `CATALOG_XLSX_PATH`：商品 Excel 路径；
- `DATABASE_PATH`：本地 SQLite 数据库路径；生产环境建议替换为托管 PostgreSQL 并实现对应 Repository；
- `ALLOWED_ORIGINS`：前端正式域名，多个域名使用逗号分隔。

不要提交 `.env`、数据库文件、API Token、真实用户档案或衣橱数据。

## GitHub Pages 品牌入口页

仓库的 `docs/` 目录包含静态品牌入口页。启用 GitHub Pages 后，项目入口地址为：

```text
https://spongebobbc018-oss.github.io/ai-wardrobe-agent/
```

在 GitHub 仓库中依次进入 **Settings → Pages**，将发布源设为 **Deploy from a branch → main → /docs** 后保存。该入口页会跳转到当前完整 Agent 服务；GitHub Pages 本身不运行 FastAPI、SQLite、推荐或日历 API。

## 恢复边界

本仓库已将原型的可恢复能力代码化，但不包含秒哒的托管运行资源或秒哒内部 Supabase 数据。要完全恢复上线服务，请按 `docs/recovery-runbook.md` 创建新运行环境、从 Excel 恢复商品，并从受管数据库备份恢复用户数据。
