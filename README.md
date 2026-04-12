# Outlook WebSocket 邮件模块

本项目当前采用双线程架构：
- 服务端线程：`aiohttp` WebSocket（邮件模块）
- 客户端线程：`aiohttp` WebSocket（内部调用方）

主入口仍为 `main.py`，按项目规则建议通过 `main.bat` 运行。

## 当前通信协议
- WebSocket 路径：`/ws/mail`
- 传输格式：`JSON-RPC 2.0`

JSON-RPC 约定：
- 请求：`{"jsonrpc":"2.0","id":1,"method":"xxx","params":{...}}`
- 成功响应：`{"jsonrpc":"2.0","id":1,"result":{...}}`
- 错误响应：`{"jsonrpc":"2.0","id":1,"error":{"code":-32000,"message":"..."}}`
- 通知（无 id）：`{"jsonrpc":"2.0","method":"xxx","params":{...}}`

## 运行流程（当前实现）
1. 服务端启动，监听 `/ws/mail`，并启动空闲检测线程（每 60 秒检查连接数）
2. 客户端连接服务端
3. 服务端先发送能力列表通知：`mail.capabilities`
4. 客户端首次请求 `auth.login`，服务端返回 `cookie`
5. 客户端发送 `auth.confirm`（携带 cookie），服务端经安全队列确认登录并返回可用接口
6. 客户端调用 `outlook.token.acquire` 请求 token 信息
7. 服务端按 cookie 维护 token（缓存到安全队列），并返回一次 `INBOX` 邮件总数（不含垃圾箱）
8. 客户端拿到首次邮件总数后，调用 `auth.logout`
9. 服务端删除 cookie 与会话信息

补充：若连接数连续两次检查为 0（即约 120 秒无人连接），服务端触发程序退出。

## 命令参数
- `--server-host`：服务端监听地址（默认 `127.0.0.1`）
- `--server-port`：服务端监听端口（默认 `8765`）
- `--client-account`：内部客户端登录账号（默认 `outlook_demo`）
- `--client-password`：内部客户端登录密码（默认 `******`）
- `--log-level`：日志等级（默认 `INFO`）
- `--log-file`：日志文件路径（默认 `log/YYYYMMDD.log`）
- `--log-retention-days`：日志保留天数（默认 `30`）

## 运行方式
```bash
cmd /c main.bat
```

PowerShell：
```powershell
powershell -Command "Invoke-Expression -Command 'cmd.exe /c main.bat'"
```

## 版本
当前版本：`26.4.12F`
最后更新：`2026-04-12`

## 更新日志
### 26.4.12F (2026-04-12)
- 重构：WebSocket 路径改为 `/ws/mail`，用于邮件模块，移除旧 `/ws/internal`
- 重构：移除全部 `/api/*` HTTP 接口，统一改为 JSON-RPC 2.0
- 新增：服务端首次连接下发 `mail.capabilities` 能力列表
- 新增：登录确认、token 维护、会话维护统一通过线程安全队列处理
- 新增：`outlook.token.acquire` 返回 token 信息及一次 `INBOX` 邮件总数
- 新增：客户端完成首次流程后主动 `auth.logout`，不再重复登录
- 新增：服务端每 60 秒检查连接数，连续两次为 0 时触发程序退出

### 26.4.12E (2026-04-12)
- 重构：`main.py` 拆分为双线程架构（服务端线程 + 客户端线程）
- 新增：`server/websocket_server.py`，服务端基于 `aiohttp` 实现内部 WebSocket 与对外 HTTP 接口
- 新增：`client/internal_ws_client.py`，客户端基于 `aiohttp` 首次登录获取 `cookie` 后周期上报账号与邮件数量
- 新增：服务端使用线程安全队列维护客户端登录信息，并提供登录会话查询接口
