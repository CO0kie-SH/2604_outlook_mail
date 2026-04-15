# Outlook WebSocket 邮件模块

🔍 自动同步 Outlook 邮件目录与标题数据

基于 `aiohttp + WebSocket + JSON-RPC 2.0` 的 Outlook 邮件采集模块。  
当前实现采用服务端/客户端双线程协同：服务端提供内部 RPC 接口，客户端执行一次性登录、同步文件夹数量与标题后主动退出。

---

## 📋 项目简介

本项目围绕 Outlook IMAP OAuth2 能力，构建了一个可本地运行的 WebSocket 邮件数据同步流程：

- 服务端负责登录会话、token 缓存、IMAP 查询与 RPC 响应
- 客户端负责按约定顺序调用 RPC，落盘 `config/*.csv`
- 支持按文件夹模式执行扩展处理：
  - `mode=num`：只同步在线邮件数
  - `mode=title`：同步数量后再拉取标题与发件人

适用于以下场景：

> 💡 推荐先用 `imap_outlook_oauth2.py --dry-run` 验证配置，再运行 `main.py`。


- 本地脚本化监控 Outlook 文件夹数量变化
- 将邮件标题按文件夹落地为 CSV，供后续自动化处理
- 作为内部模块，向其它进程提供标准化 JSON-RPC 能力

---

## ✅ 核心功能

- ✅ 自动建立本地 WebSocket 通道：`/ws/mail`
- ✅ 内置接口文档页面：`/doc/mail`、`/doc/feishu`
- ✅ 基于 JSON-RPC 2.0 的统一请求/响应格式
- ✅ 登录流程安全队列化（确认与 token 更新串行处理）
- ✅ Outlook OAuth2 token 获取（优先 refresh_token，回退设备码/缓存）
- ✅ 邮箱文件夹列表同步并落盘：`result/<前缀>_folders.csv`
- ✅ 文件夹邮件数同步：`mail.folder.count`
- ✅ 文件夹标题与发件人同步：`title`
- ✅ 支持抓取原始邮件并 URL-safe Base64 存储：`title.base64a`（字段 `Base64A`）
- ✅ 标题 CSV 自动按送达时间升序落盘（旧 -> 新）
- ✅ 抓取到标题后自动发送飞书通知（按 `config/FeiShu.csv`，单轮批量汇总）
- ✅ 文件只读保护：写入前去只读、完成后恢复只读
- ✅ 连接空闲守护：连续两次空闲检查无连接则触发进程退出
- ✅ 内置可视化页面：`/view/mail/clients`、`/view/mail/folders`、`/view/mail/titles`、`/view/mail/logout`

---

## 🗂️ 架构说明

### 线程模型

- 服务端线程：`server/websocket_server.py`
- 客户端线程：`client/internal_ws_client.py`
- 安全队列线程：服务端内部处理登录确认与 token 更新
- 空闲检测线程：每 30 秒检查活跃连接数

### 目录结构

```text
2604_outlook_mail/
├── main.py                         # 主入口（启动服务端 + 客户端线程）
├── main.bat                        # Windows 启动脚本
├── imap_outlook_oauth2.py          # Outlook 配置、OAuth2、IMAP 核心能力
├── README.md                       # 项目文档
├── 项目规则.md                      # 项目规则说明
├── client/
│   └── internal_ws_client.py       # 内部 WebSocket 客户端
├── server/
│   ├── websocket_server.py         # 内部 WebSocket 服务端
│   └── rpc_docs.py                 # 接口文档页面与渲染逻辑
├── config/
│   ├── OutLook.csv                 # Outlook 账号配置
│   └── FeiShu.csv                  # 飞书通知配置
├── result/
│   ├── *_folders.csv               # 文件夹列表与同步状态
│   └── *_<folder>.csv              # 标题结果文件
└── log/
    └── YYYYMMDD.log                # 运行日志
```

---

## 🔌 通信协议

### WebSocket 基本信息

- 路径：`/ws/mail`
- 文档：`/doc/mail`、`/doc/feishu`
- 协议：`JSON-RPC 2.0`
- 心跳：20 秒

### JSON-RPC 报文约定

- 请求

```json
{"jsonrpc":"2.0","id":1,"method":"xxx","params":{},"unixtime_ms":1710000000000}
```

- 成功响应

```json
{"jsonrpc":"2.0","id":1,"result":{},"request_unixtime_ms":1710000000000,"response_unixtime_ms":1710000000123}
```

- 错误响应

```json
{"jsonrpc":"2.0","id":1,"error":{"code":-32000,"message":"..."},"request_unixtime_ms":1710000000000,"response_unixtime_ms":1710000000123}
```

### 服务端方法列表

| 方法 | 说明 | 关键参数 |
|---|---|---|
| `auth.login` | 登录并分配 cookie | `account`, `password` |
| `auth.confirm` | 确认 cookie，返回可用方法列表 | `cookie` |
| `outlook.token.acquire` | 获取/复用 token，并返回文件夹列表 | `cookie` |
| `mail.folder.count` | 查询单文件夹邮件数量 | `cookie`, `folder_name`, `current_count` |
| `title` | 查询单文件夹邮件标题、发件人、时间 | `cookie`, `folder_name` |
| `title.base64a` | 查询单文件夹邮件标题并返回原始邮件 `Base64A` | `cookie`, `folder_name` |
| `feishu.notify` | 服务端发送飞书通知 | `cookie`, `body`, `title?`, `tag?` |
| `auth.logout` | 注销会话并清理 cookie | `cookie` |

---

## 🔄 执行流程

```text
1) 服务端启动并监听 /ws/mail
2) 客户端连接后接收 mail.capabilities
3) 客户端调用 auth.login -> 获取 cookie
4) 客户端调用 auth.confirm -> 确认登录
5) 客户端调用 outlook.token.acquire -> 获取 folders
6) 客户端落盘 result/<前缀>_folders.csv
7) 按 mode 执行扩展：
   - num   -> 调用 mail.folder.count 并回写计数
   - title -> 先 count，再调用 title 落盘标题 CSV
   - Base64A -> 先 count，再调用 title.base64a 落盘标题+原始内容 Base64
8) 客户端调用 auth.logout
9) 服务端清理会话；若长期空闲则触发进程退出
```

---

## 🔧 配置说明

### 1) Outlook 配置：`config/OutLook.csv`

字段：

- `mail`：配置名（与 `--profile` 对应）
- `user`：邮箱账号
- `password`：密码（可不使用）
- `client_id`：Azure 应用 Client ID
- `refresh_token`：刷新令牌

说明：

- CSV 使用 `utf-8-sig` 编码
- `user/password/client_id/refresh_token` 支持 URL-safe Base64（无 `=`）
- 环境变量优先级高于 CSV（如 `OUTLOOK_EMAIL`、`OUTLOOK_CLIENT_ID`）

### 2) 文件夹同步配置：`result/<前缀>_folders.csv`

字段：

- `unixtime_ms`
- `name`
- `flags`
- `mode`
- `current_count`
- `online_count`
- `current_unixtime_ms`
- `update_unixtime_ms`

`mode` 约定：

- 空：仅保留文件夹元数据
- `num`：同步邮件数量
- `title`：同步数量 + 标题数据
- `Base64A`：同步数量 + 标题数据 + 原始邮件 URL-safe Base64（字段 `Base64A`）

### 3) 标题结果文件：`result/<前缀>_<文件夹名>.csv`

字段：

- `mail_id`
- `uid`
- `message_id`
- `sender`
- `title`
- `received_at`
- `received_unixtime_ms`
- `unixtime_ms`
- `Base64A`（当 `mode=Base64A` 时有值）

### 4) 飞书配置：`config/FeiShu.csv`

当 `mode=title` 抓取到邮件标题后，客户端会发起 `feishu.notify`，由服务端统一发送飞书通知。

字段：

- `tag`：机器人标识
- `url`：飞书 Webhook 地址
- `mode`：发送模式（`none` / `text` / `title`）

说明：

- `none`：跳过该机器人
- `text`：发送纯文本
- `title`：发送带标题的富文本（post）
- 通知内容按“单轮汇总”发送，包含账号、文件夹统计和标题摘要
- 单文件夹默认展示前 5 条标题，单轮总体最多汇总 20 条，并带消息体长度截断保护

---

## 📦 环境要求

- Python：3.12+
- 依赖：
  - `aiohttp`
  - `msal`

安装示例：

```bash
pip install aiohttp msal
```

---

## 🚀 运行方式

> ⭐ 快速入口：日常建议直接执行 `cmd /c main.bat`（已内置编码与路径处理）。

### 推荐（Windows）

```bash
cmd /c main.bat
```

### PowerShell

```powershell
powershell -Command "Invoke-Expression -Command 'cmd.exe /c main.bat'"
```

### 直接运行

```bash
python main.py
```

> 📌 说明：如果你只想先验证账号配置，不启动 WebSocket 双线程，请使用 `python imap_outlook_oauth2.py --dry-run`。

---

## ⚙️ 命令参数

| 参数 | 默认值 | 说明 |
|---|---|---|
| `--server-host` | `127.0.0.1` | 服务端监听地址 |
| `--server-port` | `8765` | 服务端监听端口 |
| `--client-account` | 自动推断 | 客户端登录账号（可被 `WS_CLIENT_ACCOUNT` 覆盖） |
| `--client-password` | `******` | 客户端登录密码 |
| `--log-level` | `INFO` | 日志级别 |
| `--log-file` | `log/YYYYMMDD.log` | 日志文件 |
| `--log-retention-days` | `30` | 日志保留天数 |

---

## 🛠️ 常见操作

### 1) 首次检查配置（不连网）

```bash
python imap_outlook_oauth2.py --dry-run
```

### 2) 列出邮箱目录

```bash
python imap_outlook_oauth2.py --list-mailboxes
```

### 3) 执行完整邮件读取（直连模式）

```bash
python imap_outlook_oauth2.py --mailbox INBOX
```

### 4) 执行 WebSocket 模块一轮同步

```bash
python main.py
```

---

## 🚨 故障排查

### 快速排查表

| 现象 | 优先检查 |
|---|---|
| 启动即报配置错误 | `config/OutLook.csv`、`OUTLOOK_EMAIL`、`OUTLOOK_CLIENT_ID` |
| RPC 提示 token 未准备好 | 是否先执行 `auth.login -> auth.confirm -> outlook.token.acquire` |
| CSV 未更新 | `*_folders.csv` 的 `mode` 是否为 `num` / `title` |
| 程序无操作后退出 | 是否触发空闲守护（约 60 秒无连接） |
| 终端中文乱码 | 是否通过 `main.bat` 启动 |

- 报错 `缺少 client_id` / `缺少邮箱配置`
  - 检查 `config/OutLook.csv` 或环境变量 `OUTLOOK_EMAIL`、`OUTLOOK_CLIENT_ID`
- `token not ready, call outlook.token.acquire first`
  - 调用顺序错误，需先走 `auth.login -> auth.confirm -> outlook.token.acquire`
- 文件夹数量或标题未更新
  - 检查 `*_folders.csv` 的 `mode` 是否为 `num` 或 `title`
- 连接后很快退出
  - 服务端空闲守护生效：连续两次（约 60 秒）无连接会触发退出
- 中文乱码
  - 使用 `main.bat`（已设置 `PYTHONIOENCODING=utf-8`）

---

## 🏷️ 版本

> 🟢 Stable: `26.4.12M`

当前版本：`26.4.15D`  
最后更新：`2026-04-15`

```text
版本格式：YY.M.DX
示例：26.4.12L = 2026年4月12日 第 L 次迭代
```

---

## 📝 更新日志

### 26.4.15D (2026-04-15)

- 🔧 优化：增量标题同步避免重复读取本地标题 CSV，减少一次磁盘读开销。
- 🔧 优化：当线上数量小于本地数量时，客户端会自动裁剪本地标题 CSV，保证 `online_count` 与本地结果对齐。
- 🔧 优化：服务端在“无增量邮件”场景也输出性能日志，便于完整观察链路耗时。
- 📝 文档：`/doc/mail` 新增性能日志关键字说明（`perf title incremental`、`perf query title*`、`perf fetch title*`）。

### 26.4.15C (2026-04-15)

- 🔧 优化：增量标题同步，按 `incremental_count = max(0, online_count - local_count)` 计算增量请求量。
- 🔧 优化：`title/title.base64a` 支持可选参数 `known_max_uid`、`incremental_count`，服务端按 UID 增量查询。
- 🔧 优化：标题 CSV 改为“按 key 合并写入”而非覆盖，减少重复抓取并保留历史。
- 🔧 优化：补充增量链路性能日志（`perf title incremental`、`perf query title*`、`perf fetch title*`）。
- 📝 文档：同步 `/doc/mail` 与 README 接口说明，补齐增量参数与行为说明。

### 26.4.15B (2026-04-15)

- 🔧 优化：客户端增量标题同步流程接入 `known_max_uid` 与 `incremental_count` 参数。
- 🔧 优化：本地标题 CSV 写盘策略由覆盖改为合并（`uid` -> `message_id` -> `mail_id`）。
- 🔧 优化：新增 CSV 大字段保护（`csv.field_size_limit`），降低 `Base64A` 场景字段长度异常概率。
- 📝 文档：服务端文档补充 `title/title.base64a` 增量参数说明。

### 26.4.15A (2026-04-15)

- ✨ 新增：页面路由 `/view/mail/clients`、`/view/mail/folders`、`/view/mail/titles`、`/view/mail/logout`，支持在线会话查看、文件夹查看、标题查看和网页触发退出。
- ✨ 新增：服务端反向 RPC `mail.folders.local.list` 与 `mail.client.force.logout`，客户端支持接收并处理。
- ✨ 新增：支持 `mode=Base64A`，新增 RPC `title.base64a`，标题结果 CSV 新增字段 `Base64A`。
- 🔧 优化：`outlook.token.acquire` 改为复用 cookie 会话 IMAP 连接（断连重建），显著降低轮询耗时。
- 🔧 优化：补充性能日志（客户端端到端、服务端分段、IMAP connect/auth/list 分解）便于定位瓶颈。
- 🔧 优化：`/view/mail/titles` 改为服务端直接查询 `title`，避免读取超大 `Base64A` CSV 触发字段长度异常。
- 🔧 优化：`Base64A` 编码按项目规则改为 `urlsafe_b64encode(...).replace(b'=', b'').decode('utf-8')`。

### 26.4.12M (2026-04-12)

- 🔧 重构：将接口文档页面从 `server/websocket_server.py` 拆分到独立模块 `server/rpc_docs.py`，由独立类负责 HTML 文档渲染。
- ✨ 新增：文档路由 `/doc/mail` 与 `/doc/feishu`，覆盖参数、返回字段、错误码和 JSON-RPC 示例。
- 🔧 优化：客户端标题通知改为调用 JSON-RPC `feishu.notify`，由服务端统一执行飞书发送。
- 🔧 优化：文件夹与标题 CSV 输出路径统一为 `result/` 目录。
- ✨ 新增：标题结果字段扩展为 `mail_id,uid,message_id,sender,title,received_at,received_unixtime_ms,unixtime_ms`。

### 26.4.12L (2026-04-12)

- 🔧 优化：飞书通知由“每文件夹单独发送”改为“单轮批量汇总发送”，减少刷屏与请求次数
- 🔧 优化：新增通知摘要裁剪策略（单文件夹前 5 条、单轮最多 20 条）
- 🔧 优化：新增飞书消息体长度保护，超长内容自动截断，避免发送失败

### 26.4.12K (2026-04-12)

- ✨ 新增：迁移 `GitHub_Releases_Down` 的飞书通知能力，新增 `feishu_notifier.py`
- ✨ 新增：客户端在 `mode=title` 抓取标题并落盘后自动发送飞书消息
- 🔧 优化：飞书通知正文包含账号、文件夹、在线数量、抓取数量和最新标题列表（最多 20 条）

### 26.4.12J (2026-04-12)

- ✨ 新增：`mode=title` 返回并落盘发件人信息字段 `sender`
- 🔧 优化：标题 CSV 结构扩展为 `mail_id,title,sender,received_at,received_unixtime_ms`

### 26.4.12I (2026-04-12)

- ✨ 新增：支持 `mode=title`，客户端调用 JSON-RPC `title` 接口获取指定文件夹邮件标题
- ✨ 新增：服务端读取邮件 `Date` 送达时间并返回，客户端按送达时间降序排序
- ✨ 新增：客户端按 `前缀_文件夹名.csv` 落盘标题结果（例如 `MarkGordon7281_Inbox.csv`）

### 26.4.12H (2026-04-12)

- ✨ 新增：客户端支持读取 `*_folders.csv` 的扩展字段（`mode/current_count/online_count/current_unixtime_ms/update_unixtime_ms`）
- ✨ 新增：`mode=num` 时调用 JSON-RPC `mail.folder.count` 查询指定文件夹数量，`current_count` 为空时按 `0` 发送
- ✨ 新增：服务端增加 `mail.folder.count` 方法，返回 `success` 与对应文件夹数量
- 🔧 优化：客户端收到结果后回写本地 CSV（在线数量、当前计数、更新时间戳）

### 26.4.12G (2026-04-12)

- 🔧 优化：客户端 JSON-RPC 调用增加超时控制（20秒）与响应时间戳一致性校验
- 🔧 优化：服务端统一毫秒时间戳生成逻辑，并增强 `unixtime_ms` 非法值兜底处理
- 🔧 优化：文件夹 CSV 导出按 `name` 排序，提升结果稳定性与可比对性
- 🔧 优化：日志格式支持定位源码文件与行号，便于问题排查

### 26.4.12F (2026-04-12)

- 🔧 重构：WebSocket 路径改为 `/ws/mail`，用于邮件模块，移除旧 `/ws/internal`
- 🔧 重构：移除全部 `/api/*` HTTP 接口，统一改为 JSON-RPC 2.0
- ✨ 新增：服务端首次连接下发 `mail.capabilities` 能力列表
- ✨ 新增：登录确认、token 维护、会话维护统一通过线程安全队列处理
- ✨ 新增：`outlook.token.acquire` 返回 token 信息及一次 `INBOX` 邮件总数
- ✨ 新增：客户端完成首次流程后主动 `auth.logout`，不再重复登录
- ✨ 新增：服务端每 30 秒检查连接数，连续两次为 0 时触发程序退出

### 26.4.12E (2026-04-12)

- 🔧 重构：`main.py` 拆分为双线程架构（服务端线程 + 客户端线程）
- ✨ 新增：`server/websocket_server.py`，服务端基于 `aiohttp` 实现内部 WebSocket 与对外 HTTP 接口
- ✨ 新增：`client/internal_ws_client.py`，客户端基于 `aiohttp` 首次登录获取 `cookie` 后周期上报账号与邮件数量
- ✨ 新增：服务端使用线程安全队列维护客户端登录信息，并提供登录会话查询接口
