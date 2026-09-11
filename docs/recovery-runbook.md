# 衣橱顾问恢复手册

## 恢复目标

在秒哒、连接器或单一云平台不可用时，以 GitHub `main` 分支作为源码基线，在新环境恢复可访问的衣橱顾问服务。

## 已代码化的能力

| 模块 | 恢复来源 |
|---|---|
| 导购与搭配规则 | `agent-backend/app/main.py` |
| 移动优先网页 | `agent-backend/app/static/` |
| 商品主数据 | `AI穿搭导购_商品更新源.xlsx` |
| 表结构 | `database/migrations/0001_initial_schema.sql` |
| 本地/容器运行 | `README.md`、`Dockerfile`、`docker-compose.yml` |

## 快速恢复

1. 从 GitHub 克隆私有仓库。
2. 创建隔离 Python 虚拟环境并安装 `agent-backend/requirements.txt`。
3. 配置 `ALLOWED_ORIGINS`，本地开发可设为 `http://localhost:8080`。
4. 启动 FastAPI。首次启动会创建 SQLite 表并导入 70 条商品。
5. 访问 `/ready`，确认返回 `catalog_items: 70`。
6. 访问首页，验证输入“通勤，¥1000内，简约，165cm，M码，显瘦”能生成推荐。

## 生产恢复建议

- 使用 Cloud Run、CloudBase CloudRun、Railway 或同类容器平台运行 Docker 镜像；服务必须绑定 `0.0.0.0:$PORT`。
- 前端当前由 FastAPI 同源托管，适合单服务部署；如后续拆分前端，必须将前端正式域名加入 `ALLOWED_ORIGINS`。
- SQLite 只适用于单实例演示。多用户生产环境必须将 Repository 替换为 PostgreSQL，并把迁移纳入 CI/CD。
- 商品 Excel 是初始种子数据；真实商品、个人档案、衣橱和日历应使用受管数据库并启用每日备份。

## 数据与隐私

- 匿名 `profile_id` 在浏览器 localStorage 生成，用户可清空浏览器数据或删除档案。
- 不提交数据库文件、真实衣橱/日历记录、Token 或 `.env`。
- 身高、尺码、舒适度仅处理用户主动提供的数据，并且仅影响软排序或尺码提示。

## 发布前检查

1. `python -m py_compile agent-backend/app/main.py`
2. `GET /health` 返回 `status=ok`
3. `GET /ready` 返回 `catalog_items=70`
4. 推荐、档案、衣橱新增与日历同日更新均通过冒烟测试
5. 使用独立浏览器会话，确认匿名档案不共享
6. 使用未提交密钥的干净环境进行容器构建
