# Outlook IMAP 邮件读取工具

一个基于 OAuth2 的 Outlook IMAP 邮件读取脚本，支持使用 `refresh_token` 或设备码流程登录，读取邮箱中的邮件标题与正文。

## 功能
- OAuth2 登录（优先 `refresh_token`，兜底设备码）
- 连接 Outlook IMAP（`outlook.office365.com:993`）
- 读取指定邮箱目录（默认 `INBOX`）
- 输出邮件总数、每封邮件标题与正文
- 兼容 `text/plain` 与 `text/html`（HTML 自动转文本）

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

## 运行方式
按项目规则推荐：
```bash
cmd /c main.bat
```

PowerShell：
```powershell
powershell -Command "Invoke-Expression -Command 'cmd.exe /c main.bat'"
```

## 环境变量
- `OUTLOOK_EMAIL`：邮箱地址
- `OUTLOOK_CLIENT_ID`：Azure 应用 Client ID
- `OUTLOOK_REFRESH_TOKEN`：可选，存在时优先使用
- `OUTLOOK_TENANT`：可选，默认 `consumers`
- `OUTLOOK_IMAP_HOST`：可选，默认 `outlook.office365.com`
- `OUTLOOK_IMAP_PORT`：可选，默认 `993`
- `OUTLOOK_IMAP_MAILBOX`：可选，默认 `INBOX`
- `OUTLOOK_TOKEN_CACHE`：可选，默认 `.outlook_token_cache.json`

## 示例
```powershell
$env:OUTLOOK_EMAIL='***'
$env:OUTLOOK_CLIENT_ID='***'
$env:OUTLOOK_REFRESH_TOKEN='***'
cmd /c main.bat
```

## 版本
当前版本：`26.4.11A`
最后更新：`2026-04-11`

## 更新日志
### 26.4.11A (2026-04-11)
- 新增：`main.py` 作为统一入口
- 新增：`main.bat` 作为 Windows 启动脚本，并设置 `PYTHONIOENCODING=utf-8`
- 新增：`tmp`、`db`、`config`、`log` 目录
- 文档：按项目规则重写 `README.md`

