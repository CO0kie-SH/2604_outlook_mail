# Outlook IMAP 邮件读取工具

一个基于 OAuth2 的 Outlook IMAP 邮件读取脚本，支持使用 `refresh_token` 或设备码流程登录，读取邮箱中的邮件标题与正文。

## 功能
- OAuth2 登录（优先 `refresh_token`，兜底设备码）
- 连接 Outlook IMAP（`outlook.office365.com:993`）
- 读取指定邮箱目录（默认 `INBOX`）
- 输出邮件总数、每封邮件时间/标题/正文
- 兼容 `text/plain` 与 `text/html`（HTML 自动转文本）
- 支持从 `config/OutLook.csv` 读取账号配置
- 支持日志等级、按天日志文件与过期日志清理
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
2. CSV 文件（优先读取 `config/OutLook.local.csv`，不存在时读取 `config/OutLook.csv`）

### CSV 配置格式
推荐文件：
- `config/OutLook.local.csv`（本地私有，不提交）
- `config/OutLook.csv`（仓库示例）

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
- `--list-mailboxes`：列出邮箱目录，不读取邮件内容
- `--mailbox`：邮箱目录（默认 `INBOX`）
- `--profile`：CSV 中 `mail` 字段（默认 `outlook`）
- `--config`：CSV 配置路径（默认：若存在 `config/OutLook.local.csv` 则优先使用，否则使用 `config/OutLook.csv`）
- `--log-level`：日志等级（默认 `INFO`）
- `--log-file`：日志文件路径（默认 `log/YYYYMMDD.log`）
- `--log-retention-days`：日志保留天数（默认 `30`）

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
- `OUTLOOK_CONFIG_PATH`：可选；未设置时默认自动选择 `config/OutLook.local.csv` 或 `config/OutLook.csv`
- `OUTLOOK_LOG_LEVEL`：可选，默认 `INFO`
- `OUTLOOK_LOG_FILE`：可选，默认 `log/YYYYMMDD.log`
- `OUTLOOK_LOG_RETENTION_DAYS`：可选，默认 `30`

## 版本
当前版本：`26.4.12D`
最后更新：`2026-04-12`

## 更新日志
### 26.4.12D (2026-04-12)
- 新增：默认日志路径改为 `log/YYYYMMDD.log`，并支持 `--log-retention-days` 自动清理过期日志
- 新增：`main.py` 统一初始化 logger，并在启动时打印运行目录、项目目录、日志目录、解释器路径
- 新增：邮件输出与日志记录均包含时间戳字段（基于邮件 `Date` 头）
- 优化：提取单封邮件解析逻辑，减少 `fetch_all_mails` 内重复分支代码，提升可维护性

### 26.4.12C (2026-04-12)
- 新增：`--list-mailboxes` 参数，可直接列出 IMAP 目录与 flags
- 优化：邮箱目录选择支持常见别名映射（如 `junk`、`垃圾邮件` 自动匹配 `\\Junk`）
- 优化：当目录选择失败时，报错中附带可用目录列表，便于排查

### 26.4.12B (2026-04-12)
- 安全：仓库内 `config/OutLook.csv` 改为脱敏示例配置，避免提交真实凭证
- 新增：支持默认优先读取 `config/OutLook.local.csv`（用于本地私有配置）
- 新增：`.gitignore` 忽略本地 token 缓存、日志、临时目录与本地私有配置
- 优化：`--dry-run` 输出默认脱敏展示 `email` 与 `client_id`

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
