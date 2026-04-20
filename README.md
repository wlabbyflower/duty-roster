# Duty Roster

值班表展示与企业微信机器人通知工具。

## 功能

- 首页仅展示今日值班：售前是谁、售后是谁
- 右上角设置面板可进行：
  - 日常值班 Excel 导入
  - 节假日值班 Excel 导入
  - 在线编辑排班并保存
  - 导出 Excel
  - 企业微信机器人通知配置
- 企业微信机器人支持：
  - Webhook 配置
  - 每天通知时间（HH:MM）
  - 通知人数（按 UserID 顺序取前 N 个 @）
  - 时区（默认 `Asia/Shanghai`）
  - 测试通知
- 通知内容支持按人展示状态：如 `今日休息` / `下午来` / `全天都在`
- 当天命中节假日排班时，通知仅发送节假日值班内容；未命中时按日常值班正常通知

## Excel 导入

支持两类导入：

1. 日常值班 Excel

支持两类模板：

1. 三列表（推荐）
- `日期`
- `售前`
- `售后`

2. 矩阵模板
- 表头包含 `日期`、`时间`，后续为人员列
- 单元格中包含 `售前`/`售后` 关键字即可聚合
- 可识别按日期排班和按周模板（周一~周日）

2. 节假日值班 Excel
- 仍支持矩阵排班内容解析
- 需要在表格顶部写明节假日起止日期，例如：
  - `2026年4月20日-2026年4月22日`
  - `4月20日到4月22日`
- 系统会把节假日排班单独存储，并在节假日期间优先覆盖日常排班通知

## 目录结构

```text
apps/duty-roster/
├─ docker-compose.yml
├─ README.md
└─ backend/
   ├─ main.py
   ├─ Dockerfile
   ├─ requirements.txt
   ├─ data/
   │  ├─ .gitkeep
   │  └─ .gitignore
   └─ static/
      ├─ index.html
      ├─ style.css
      └─ app.js
```

## 本地运行

```bash
cd apps/duty-roster/backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 8000
```

访问：`http://localhost:8000`

## Docker Compose 部署（推荐）

```bash
cd apps/duty-roster
docker compose up -d --build
```

访问：`http://<服务器IP>:8000`

## 开机自启

`docker-compose.yml` 已配置：

- `restart: unless-stopped`

只要 Docker 服务本身是开机自启，容器会自动拉起。可用以下命令确认：

```bash
sudo systemctl enable docker
sudo systemctl status docker
```

## 更新部署

```bash
cd apps/duty-roster
docker compose up -d --build
```
