# Outlook IMAP 邮件读取工具

一个基于 OAuth2 的 Outlook IMAP 邮件读取脚本，支持使用 `refresh_token` 或设备码流程登录，读取邮箱中的邮件标题与正文。

## 功能
- OAuth2 登录（优先 `refresh_token`，兜底设备码）
- 连接 Outlook IMAP（`outlook.office365.com:993`）
- 读取指定邮箱目录（默认 `INBOX`）
- 输出邮件总数、每封邮件标题与正文
- 兼容 `text/plain` 与 `text/html`（HTML 自动转文本）
- 支持从 `config/OutLook.csv` 读取账号配置
- 支持日志等级与日志文件输出
- 代码采用类结构（`OutlookMailService`）

## 项目结构
```text
2604_outlook_mail/
├── main.py
├── main.bat
├── imap_outlook_oauth2.py
├── 项目规则.md
├── README.md
├── config/
├── db/
├── log/
└── tmp/
```

## 环境要求
- Python 3.12
- 推荐解释器：`D:\0Code2\py312\python.exe`
- 依赖：`msal`、`aiohttp`

安装依赖：
```bash
D:\0Code2\py312\python.exe -m pip install msal aiohttp
```

## 配置方式
支持两种方式，环境变量优先级更高：
1. 环境变量（可覆盖 CSV）
2. `config/OutLook.csv`

### CSV 配置格式
文件：`config/OutLook.csv`

表头：
```csv
mail,user,password,client_id,refresh_token
```

说明：
- `mail`：配置标识（默认读取 `outlook`）
- `user`：邮箱地址
- `client_id`：Azure 应用 Client ID
- `refresh_token`：可选，存在时优先使用
- 字段可为明文或 URL-safe Base64（代码内自动尝试解码）

## 运行方式
按项目规则推荐：
```bash
cmd /c main.bat
```

PowerShell：
```powershell
powershell -Command "Invoke-Expression -Command 'cmd.exe /c main.bat'"
```

直接运行：
```powershell
$env:PYTHONIOENCODING='utf-8'
D:\0Code2\py312\python.exe -B main.py
```

## 命令参数
- `--dry-run`：只检查配置，不连接网络
- `--mailbox`：邮箱目录（默认 `INBOX`）
- `--profile`：CSV 中 `mail` 字段（默认 `outlook`）
- `--config`：CSV 配置路径（默认 `config/OutLook.csv`）
- `--log-level`：日志等级（默认 `INFO`）
- `--log-file`：日志文件路径（默认空，仅控制台）

示例：
```powershell
$env:PYTHONIOENCODING='utf-8'
D:\0Code2\py312\python.exe -B main.py --profile outlook --log-level INFO --log-file log\outlook.log
```

## 环境变量
- `OUTLOOK_EMAIL`：邮箱地址（可覆盖 CSV）
- `OUTLOOK_CLIENT_ID`：Azure 应用 Client ID（可覆盖 CSV）
- `OUTLOOK_REFRESH_TOKEN`：可选，存在时优先使用（可覆盖 CSV）
- `OUTLOOK_TENANT`：可选，默认 `consumers`
- `OUTLOOK_IMAP_HOST`：可选，默认 `outlook.office365.com`
- `OUTLOOK_IMAP_PORT`：可选，默认 `993`
- `OUTLOOK_IMAP_MAILBOX`：可选，默认 `INBOX`
- `OUTLOOK_TOKEN_CACHE`：可选，默认 `.outlook_token_cache.json`
- `OUTLOOK_PROFILE`：可选，默认 `outlook`
- `OUTLOOK_CONFIG_PATH`：可选，默认 `config/OutLook.csv`
- `OUTLOOK_LOG_LEVEL`：可选，默认 `INFO`
- `OUTLOOK_LOG_FILE`：可选，默认空

## 版本
当前版本：`26.4.12A`
最后更新：`2026-04-12`

## 更新日志
### 26.4.12A (2026-04-12)
- 重构：核心逻辑改为类结构，新增 `OutlookConfig` 与 `OutlookMailService`
- 新增：从 `config/OutLook.csv` 自动读取邮箱配置，并支持按 `profile` 选择
- 新增：CSV 字段自动尝试 URL-safe Base64 解码（兼容明文）
- 新增：日志系统（控制台 + 可选文件输出），支持 `--log-level`、`--log-file`
- 文档：同步更新运行参数、配置说明、示例

### 26.4.11A (2026-04-11)
- 新增：`main.py` 作为统一入口
- 新增：`main.bat` 作为 Windows 启动脚本，并设置 `PYTHONIOENCODING=utf-8`
- 新增：`tmp`、`db`、`config`、`log` 目录
- 文档：按项目规则重写 `README.md`
