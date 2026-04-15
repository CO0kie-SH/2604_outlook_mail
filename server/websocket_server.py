import asyncio
import base64
import email
import imaplib
import json
import os
import re
import secrets
import threading
import time
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime, parseaddr
from html import escape
from queue import Empty, Queue
from types import SimpleNamespace
from typing import Any
from urllib.parse import quote

from aiohttp import web

from feishu_notifier import send_feishu_message
import imap_outlook_oauth2
from server.rpc_docs import InternalWSDocPages


class InternalWSServer:
    """邮件模块 WebSocket 服务端，使用 JSON-RPC 2.0 进行内部通信。"""

    IDLE_CHECK_INTERVAL_SECONDS = 30
    IDLE_ZERO_LIMIT = 2
    QUEUE_ACK_TIMEOUT_SECONDS = 3
    IMAP_KEEPALIVE_INTERVAL_SECONDS = 60

    RPC_METHOD_LOGIN = "auth.login"
    RPC_METHOD_CONFIRM = "auth.confirm"
    RPC_METHOD_ACQUIRE = "outlook.token.acquire"
    RPC_METHOD_FOLDER_COUNT = "mail.folder.count"
    RPC_METHOD_TITLE = "title"
    RPC_METHOD_TITLE_BASE64A = "title.base64a"
    RPC_METHOD_LOCAL_FOLDER_LIST = "mail.folders.local.list"
    RPC_METHOD_LOCAL_TITLE_LIST = "mail.titles.local.list"
    RPC_METHOD_CLIENT_FORCE_LOGOUT = "mail.client.force.logout"
    RPC_METHOD_FEISHU_NOTIFY = "feishu.notify"
    RPC_METHOD_LOGOUT = "auth.logout"

    def __init__(self, host: str, port: int, logger):
        self.host = host
        self.port = port
        self.logger = logger
        self._doc_pages = InternalWSDocPages()

        self.ready_event = threading.Event()
        self.shutdown_requested_event = threading.Event()
        self._stop_event = threading.Event()

        self._thread: threading.Thread | None = None
        self._consumer_thread: threading.Thread | None = None
        self._idle_checker_thread: threading.Thread | None = None
        self._imap_keepalive_thread: threading.Thread | None = None

        self._loop: asyncio.AbstractEventLoop | None = None
        self._runner: web.AppRunner | None = None

        # 安全队列：串行处理登录确认和 token 缓存等敏感会话动作。
        self._secure_queue: Queue[dict[str, Any]] = Queue()
        self._sessions: dict[str, dict[str, Any]] = {}
        self._session_lock = threading.Lock()

        self._clients_lock = threading.Lock()
        self._active_clients = 0
        self._server_to_client_rpc_id = 1_000_000
        self._server_to_client_pending: dict[int, tuple[web.WebSocketResponse, asyncio.Future]] = {}

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return

        self._stop_event.clear()
        self._consumer_thread = threading.Thread(target=self._consume_secure_queue, name="ws-secure-queue", daemon=True)
        self._consumer_thread.start()

        # 每 30 秒检查一次客户端连接数，连续两次为 0 则退出。
        self._idle_checker_thread = threading.Thread(target=self._idle_exit_checker, name="ws-idle-checker", daemon=True)
        self._idle_checker_thread.start()

        # 后台 IMAP 保活：定期对已建立的会话连接执行 NOOP，降低连接被服务端回收概率。
        self._imap_keepalive_thread = threading.Thread(
            target=self._imap_keepalive_loop,
            name="ws-imap-keepalive",
            daemon=True,
        )
        self._imap_keepalive_thread.start()

        self._thread = threading.Thread(target=self._run_server_thread, name="ws-mail-server", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        if self._loop and self._loop.is_running():
            fut = asyncio.run_coroutine_threadsafe(self._shutdown(), self._loop)
            try:
                fut.result(timeout=5)
            except Exception:
                self.logger.exception("server shutdown coroutine failed")
            self._loop.call_soon_threadsafe(self._loop.stop)

        if self._thread:
            self._thread.join(timeout=5)
        if self._consumer_thread:
            self._consumer_thread.join(timeout=5)
        if self._idle_checker_thread:
            self._idle_checker_thread.join(timeout=5)
        if self._imap_keepalive_thread:
            self._imap_keepalive_thread.join(timeout=5)

    def _run_server_thread(self) -> None:
        self._loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self._loop)
        self._loop.run_until_complete(self._start_app())
        self.logger.info("mail websocket server started on %s:%s", self.host, self.port)
        self.ready_event.set()
        self._loop.run_forever()
        self._loop.close()

    async def _start_app(self) -> None:
        app = web.Application()
        app.add_routes(
            [
                web.get("/ws/mail", self._handle_mail_ws),
                web.get("/doc/mail", self._doc_pages.handle_doc_mail),
                web.get("/doc/feishu", self._doc_pages.handle_doc_feishu),
                web.get("/view/mail/folders", self._handle_view_mail_folders),
                web.get("/view/mail/titles", self._handle_view_mail_titles),
                web.get("/view/mail/clients", self._handle_view_mail_clients),
                web.get("/view/mail/logout", self._handle_view_mail_logout),
            ]
        )
        self._runner = web.AppRunner(app)
        await self._runner.setup()
        site = web.TCPSite(self._runner, self.host, self.port)
        await site.start()
        if self.host in {"127.0.0.1", "localhost"}:
            self.logger.warning(
                "server host is loopback (%s), web pages are only reachable from this machine. "
                "For external access, run with --server-host 0.0.0.0",
                self.host,
            )

    async def _shutdown(self) -> None:
        if self._runner is not None:
            await self._runner.cleanup()
            self._runner = None

    def _idle_exit_checker(self) -> None:
        zero_count = 0
        while not self._stop_event.is_set():
            time.sleep(self.IDLE_CHECK_INTERVAL_SECONDS)
            with self._clients_lock:
                active = self._active_clients

            if active == 0:
                zero_count += 1
                self.logger.info("idle-check: active_clients=0, zero_count=%s", zero_count)
                if zero_count >= self.IDLE_ZERO_LIMIT:
                    self.logger.warning("active_clients is 0 for two checks, requesting shutdown")
                    self.shutdown_requested_event.set()
                    return
            else:
                zero_count = 0
                self.logger.info("idle-check: active_clients=%s, keep running", active)

    def _consume_secure_queue(self) -> None:
        while not self._stop_event.is_set() or not self._secure_queue.empty():
            try:
                event = self._secure_queue.get(timeout=0.5)
            except Empty:
                continue

            event_type = str(event.get("type", "")).strip()
            cookie = str(event.get("cookie", "")).strip()
            ack_event = event.get("ack")

            try:
                if event_type == "login_confirm":
                    with self._session_lock:
                        session = self._sessions.get(cookie)
                        if session:
                            session["confirmed"] = True
                            session["confirmed_at"] = datetime.now(timezone.utc).isoformat()

                elif event_type == "token_update":
                    with self._session_lock:
                        session = self._sessions.get(cookie)
                        if session:
                            session["access_token"] = str(event.get("access_token", ""))
                            session["token_acquired_at"] = datetime.now(timezone.utc).isoformat()
                            session["folders"] = event.get("folders", [])
                else:
                    self.logger.warning("unknown secure queue event: %s", event_type)
            finally:
                if isinstance(ack_event, threading.Event):
                    ack_event.set()

    def _imap_keepalive_loop(self) -> None:
        while not self._stop_event.is_set():
            time.sleep(self.IMAP_KEEPALIVE_INTERVAL_SECONDS)
            with self._session_lock:
                sessions = list(self._sessions.values())

            for session in sessions:
                imap = session.get("imap_client")
                if imap is None:
                    continue
                lock = session.get("imap_lock")
                if not hasattr(lock, "acquire") or not hasattr(lock, "release"):
                    continue
                if not lock.acquire(timeout=1):
                    continue
                try:
                    imap = session.get("imap_client")
                    if imap is None:
                        continue
                    typ, _ = imap.noop()
                    if typ != "OK":
                        self.logger.warning("imap keepalive failed: typ=%s, drop session connection", typ)
                        self._drop_session_imap(session)
                except (imaplib.IMAP4.abort, imaplib.IMAP4.error, OSError, EOFError):
                    self.logger.warning("imap keepalive connection error, drop session connection", exc_info=True)
                    self._drop_session_imap(session)
                except Exception:
                    self.logger.warning("imap keepalive unexpected error", exc_info=True)
                    self._drop_session_imap(session)
                finally:
                    lock.release()

    async def _handle_mail_ws(self, request: web.Request) -> web.StreamResponse:
        ws = web.WebSocketResponse(heartbeat=20)
        await ws.prepare(request)
        self._mark_client_connected()
        self.logger.info("mail websocket client connected")

        current_cookie = ""
        try:
            await ws.send_json(
                self._rpc_notification(
                    method="mail.capabilities",
                    params={
                        "module": "mail",
                        "methods": [
                            self.RPC_METHOD_LOGIN,
                            self.RPC_METHOD_CONFIRM,
                            self.RPC_METHOD_ACQUIRE,
                            self.RPC_METHOD_FOLDER_COUNT,
                            self.RPC_METHOD_TITLE,
                            self.RPC_METHOD_TITLE_BASE64A,
                            self.RPC_METHOD_FEISHU_NOTIFY,
                            self.RPC_METHOD_LOGOUT,
                        ],
                    },
                )
            )

            async for msg in ws:
                if msg.type != web.WSMsgType.TEXT:
                    if msg.type in {web.WSMsgType.CLOSE, web.WSMsgType.CLOSING, web.WSMsgType.CLOSED}:
                        break
                    continue

                if self._try_resolve_server_to_client_rpc_response(ws, msg.data):
                    continue

                response, new_cookie = await self._dispatch_rpc_text(msg.data, current_cookie)
                if new_cookie is not None:
                    current_cookie = new_cookie
                    if current_cookie:
                        with self._session_lock:
                            session = self._sessions.get(current_cookie)
                            if session is not None:
                                session["ws"] = ws
                if response is not None:
                    await ws.send_json(response)
        finally:
            self._cleanup_server_to_client_pending_by_ws(ws)
            if current_cookie:
                self._delete_cookie_session(current_cookie, reason="ws_disconnected")
            self._mark_client_disconnected()
            self.logger.info("mail websocket client disconnected")

        return ws

    async def _handle_view_mail_folders(self, request: web.Request) -> web.Response:
        self.logger.info(
            "view folders handler hit: path=%s remote=%s query=%s",
            request.path_qs,
            request.remote,
            dict(request.query),
        )
        cookie = str(request.query.get("cookie", "")).strip()
        if not cookie:
            self.logger.warning("view folders rejected: missing cookie")
            return web.Response(
                text=(
                    "<h3>cookie required</h3>"
                    "<p>use: <code>/view/mail/folders?cookie=YOUR_COOKIE</code></p>"
                ),
                content_type="text/html",
                charset="utf-8",
                status=400,
            )

        with self._session_lock:
            session = self._sessions.get(cookie)

        if session is None:
            self.logger.warning("view folders rejected: cookie not found, cookie=%s", cookie)
            return web.Response(
                text=(
                    "<h3>不存在的客户端</h3>"
                    "<p>cookie 对应的客户端会话不存在，或已登出。</p>"
                ),
                content_type="text/html",
                charset="utf-8",
                status=404,
            )

        ws = session.get("ws")
        account_hint = str(session.get("account", "")).strip() or "(unknown)"
        self.logger.info(
            "view folders session found: cookie=%s account=%s ws_exists=%s ws_closed=%s",
            cookie,
            account_hint,
            isinstance(ws, web.WebSocketResponse),
            bool(ws.closed) if isinstance(ws, web.WebSocketResponse) else True,
        )
        if not isinstance(ws, web.WebSocketResponse) or ws.closed:
            self.logger.warning(
                "view folders rejected: client offline, cookie=%s account=%s",
                cookie,
                account_hint,
            )
            return web.Response(
                text=(
                    "<h3>不存在的客户端</h3>"
                    "<p>cookie 存在，但客户端当前不在线。</p>"
                ),
                content_type="text/html",
                charset="utf-8",
                status=404,
            )

        ok, data = await self._call_client_rpc(
            ws,
            method=self.RPC_METHOD_LOCAL_FOLDER_LIST,
            params={"cookie": cookie},
            timeout_seconds=8,
        )
        self.logger.info(
            "view folders request: mail_type=%s account=%s url=http://%s:%s/view/mail/folders?cookie=%s",
            "outlook",
            account_hint,
            self.host,
            self.port,
            cookie,
        )
        if not ok:
            message = escape(str(data.get("message", "client rpc failed")))
            self.logger.warning(
                "view folders failed: cookie=%s account=%s message=%s",
                cookie,
                account_hint,
                message,
            )
            return web.Response(
                text=(
                    "<h3>客户端查询失败</h3>"
                    f"<p>{message}</p>"
                ),
                content_type="text/html",
                charset="utf-8",
                status=502,
            )

        account = str(data.get("account", "")).strip() or str(session.get("account", "")).strip() or "(unknown)"
        login_at = str(session.get("login_at", "")).strip() or ""
        confirmed = bool(session.get("confirmed", False))
        token_ready = bool(str(session.get("access_token", "")).strip())
        csv_path = str(data.get("csv_path", "")).strip()
        folders = data.get("folders", [])
        folders = folders if isinstance(folders, list) else []
        self.logger.info(
            "view folders client rpc success: cookie=%s account=%s csv_path=%s folder_items=%s",
            cookie,
            account,
            csv_path,
            len(folders),
        )
        folder_rows: list[str] = []
        for idx, item in enumerate(folders, start=1):
            if not isinstance(item, dict):
                continue
            raw_name = str(item.get("name", "")).strip() or "(empty)"
            name = escape(raw_name)
            flags = escape(str(item.get("flags", "")).strip() or "-")
            mode = escape(str(item.get("mode", "")).strip() or "-")
            current_count = escape(str(item.get("current_count", "")).strip() or "-")
            online_count = escape(str(item.get("online_count", "")).strip() or "-")
            title_link = f"/view/mail/titles?cookie={quote(cookie, safe='')}&folder_name={quote(raw_name, safe='')}"
            folder_rows.append(
                f"<tr><td>{idx}</td><td><a href=\"{title_link}\">{name}</a></td><td>{flags}</td><td>{mode}</td><td>{current_count}</td><td>{online_count}</td></tr>"
            )

        rows_html = "".join(folder_rows) if folder_rows else "<tr><td colspan='6'>(no folders)</td></tr>"
        html = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Mail Folders</title>
  <style>
    body {{ font-family: "Segoe UI", sans-serif; margin: 20px; color: #0f172a; background: #f8fafc; }}
    .card {{ background: #fff; border: 1px solid #dbe2ea; border-radius: 10px; padding: 14px; max-width: 980px; }}
    h1 {{ margin: 0 0 12px; font-size: 22px; }}
    .meta {{ color: #334155; font-size: 14px; line-height: 1.7; margin-bottom: 12px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
    th, td {{ border: 1px solid #e2e8f0; padding: 8px 10px; text-align: left; }}
    th {{ background: #f1f5f9; }}
    .warn {{ margin-top: 10px; color: #b45309; font-size: 13px; }}
  </style>
</head>
<body>
  <div class="card">
    <h1>邮箱文件夹列表</h1>
    <div class="meta">
      account: <b>{escape(account)}</b><br/>
      cookie: <code>{escape(cookie)}</code><br/>
      login_at: <code>{escape(login_at or "-")}</code><br/>
      confirmed: <b>{str(confirmed).lower()}</b><br/>
      token_ready: <b>{str(token_ready).lower()}</b><br/>
      csv_path: <code>{escape(csv_path or "-")}</code><br/>
      folder_count: <b>{len(folder_rows)}</b>
    </div>
    <table>
      <thead>
        <tr><th>#</th><th>name</th><th>flags</th><th>mode</th><th>current_count</th><th>online_count</th></tr>
      </thead>
      <tbody>
        {rows_html}
      </tbody>
    </table>
    {"<div class='warn'>客户端本地 CSV 没有可用数据。</div>" if not folder_rows else ""}
  </div>
</body>
</html>"""
        return web.Response(text=html, content_type="text/html", charset="utf-8")

    async def _handle_view_mail_titles(self, request: web.Request) -> web.Response:
        cookie = str(request.query.get("cookie", "")).strip()
        folder_name = str(request.query.get("folder_name", "")).strip()
        if not cookie or not folder_name:
            return web.Response(
                text=(
                    "<h3>cookie and folder_name required</h3>"
                    "<p>use: <code>/view/mail/titles?cookie=YOUR_COOKIE&folder_name=Inbox</code></p>"
                ),
                content_type="text/html",
                charset="utf-8",
                status=400,
            )

        with self._session_lock:
            session = self._sessions.get(cookie)
        if session is None:
            return web.Response(
                text="<h3>不存在的客户端</h3><p>cookie 对应的客户端会话不存在，或已登出。</p>",
                content_type="text/html",
                charset="utf-8",
                status=404,
            )

        ws = session.get("ws")
        account_hint = str(session.get("account", "")).strip() or "(unknown)"
        if not isinstance(ws, web.WebSocketResponse) or ws.closed:
            return web.Response(
                text="<h3>不存在的客户端</h3><p>cookie 存在，但客户端当前不在线。</p>",
                content_type="text/html",
                charset="utf-8",
                status=404,
            )

        ok, data = await asyncio.to_thread(self._query_folder_titles, cookie, folder_name)
        if not ok:
            message = escape(str(data.get("message", "client rpc failed")))
            return web.Response(
                text=(
                    "<h3>客户端查询失败</h3>"
                    f"<p>{message}</p>"
                ),
                content_type="text/html",
                charset="utf-8",
                status=502,
            )

        account = account_hint
        titles = data.get("titles", [])
        titles = titles if isinstance(titles, list) else []
        title_rows: list[str] = []
        for idx, item in enumerate(titles, start=1):
            if not isinstance(item, dict):
                continue
            mail_id = escape(str(item.get("mail_id", "")).strip() or "-")
            uid = escape(str(item.get("uid", "")).strip() or "-")
            message_id = escape(str(item.get("message_id", "")).strip() or "-")
            sender = escape(str(item.get("sender", "")).strip() or "(未知发件人)")
            title = escape(str(item.get("title", "")).strip() or "(无主题)")
            received_at = escape(str(item.get("received_at", "")).strip() or "-")
            title_rows.append(
                "<tr>"
                f"<td>{idx}</td><td>{mail_id}</td><td>{uid}</td><td>{message_id}</td><td>{sender}</td><td>{title}</td><td>{received_at}</td>"
                "</tr>"
            )

        rows_html = "".join(title_rows) if title_rows else "<tr><td colspan='7'>(no titles)</td></tr>"
        html = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Mail Titles</title>
  <style>
    body {{ font-family: "Segoe UI", sans-serif; margin: 20px; color: #0f172a; background: #f8fafc; }}
    .card {{ background: #fff; border: 1px solid #dbe2ea; border-radius: 10px; padding: 14px; max-width: 1300px; }}
    h1 {{ margin: 0 0 12px; font-size: 22px; }}
    .meta {{ color: #334155; font-size: 14px; line-height: 1.7; margin-bottom: 12px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
    th, td {{ border: 1px solid #e2e8f0; padding: 8px 10px; text-align: left; vertical-align: top; }}
    th {{ background: #f1f5f9; }}
    code {{ font-family: Consolas, "Courier New", monospace; font-size: 12px; }}
  </style>
</head>
<body>
  <div class="card">
    <h1>文件夹标题列表</h1>
    <div class="meta">
      account: <b>{escape(account)}</b><br/>
      cookie: <code>{escape(cookie)}</code><br/>
      folder_name: <b>{escape(folder_name)}</b><br/>
      title_count: <b>{len(title_rows)}</b><br/>
      <a href="/view/mail/folders?cookie={quote(cookie, safe='')}">返回文件夹列表</a>
    </div>
    <table>
      <thead>
        <tr><th>#</th><th>mail_id</th><th>uid</th><th>message_id</th><th>sender</th><th>title</th><th>received_at</th></tr>
      </thead>
      <tbody>{rows_html}</tbody>
    </table>
  </div>
</body>
</html>"""
        return web.Response(text=html, content_type="text/html", charset="utf-8")

    async def _handle_view_mail_clients(self, request: web.Request) -> web.Response:
        with self._session_lock:
            sessions = [dict(item) for item in self._sessions.values()]

        online_sessions: list[dict[str, Any]] = []
        for session in sessions:
            ws = session.get("ws")
            if isinstance(ws, web.WebSocketResponse) and not ws.closed:
                online_sessions.append(session)

        rows: list[str] = []
        for idx, session in enumerate(online_sessions, start=1):
            account = escape(str(session.get("account", "")).strip() or "(unknown)")
            cookie_raw = str(session.get("cookie", "")).strip()
            cookie = escape(cookie_raw or "-")
            login_at = escape(str(session.get("login_at", "")).strip() or "-")
            last_query_at = escape(str(session.get("last_query_at", "")).strip() or "-")
            folder_count = len(session.get("folders", [])) if isinstance(session.get("folders", []), list) else 0
            folder_link = f"/view/mail/folders?cookie={quote(cookie_raw, safe='')}" if cookie_raw else "#"
            logout_link = f"/view/mail/logout?cookie={quote(cookie_raw, safe='')}" if cookie_raw else "#"
            rows.append(
                "<tr>"
                f"<td>{idx}</td>"
                f"<td>{account}</td>"
                f"<td><code>{cookie}</code></td>"
                f"<td><code>{login_at}</code></td>"
                f"<td><code>{last_query_at}</code></td>"
                f"<td><a href=\"{folder_link}\">{folder_count}</a></td>"
                f"<td><a href=\"{logout_link}\">logout</a></td>"
                "</tr>"
            )

        rows_html = "".join(rows) if rows else "<tr><td colspan='7'>(no online clients)</td></tr>"
        html = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Mail Clients</title>
  <style>
    body {{ font-family: "Segoe UI", sans-serif; margin: 20px; color: #0f172a; background: #f8fafc; }}
    .card {{ background: #fff; border: 1px solid #dbe2ea; border-radius: 10px; padding: 14px; max-width: 1180px; }}
    h1 {{ margin: 0 0 10px; font-size: 22px; }}
    .meta {{ color: #334155; font-size: 14px; margin-bottom: 10px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
    th, td {{ border: 1px solid #e2e8f0; padding: 8px 10px; text-align: left; }}
    th {{ background: #f1f5f9; }}
    code {{ font-family: Consolas, "Courier New", monospace; font-size: 12px; }}
  </style>
</head>
<body>
  <div class="card">
    <h1>当前服务端连接的客户端</h1>
    <div class="meta">online_clients: <b>{len(online_sessions)}</b></div>
    <table>
      <thead>
        <tr><th>#</th><th>邮箱全名</th><th>cookie</th><th>登录时间</th><th>最后一次查询时间</th><th>文件夹数量</th><th>退出</th></tr>
      </thead>
      <tbody>{rows_html}</tbody>
    </table>
  </div>
</body>
</html>"""
        return web.Response(text=html, content_type="text/html", charset="utf-8")

    async def _handle_view_mail_logout(self, request: web.Request) -> web.Response:
        cookie = str(request.query.get("cookie", "")).strip()
        if not cookie:
            return web.Response(
                text=(
                    "<h3>cookie required</h3>"
                    "<p>use: <code>/view/mail/logout?cookie=YOUR_COOKIE</code></p>"
                ),
                content_type="text/html",
                charset="utf-8",
                status=400,
            )

        with self._session_lock:
            session = self._sessions.get(cookie)
        if session is None:
            return web.Response(
                text="<h3>不存在的客户端</h3><p>cookie 对应的客户端会话不存在，或已登出。</p>",
                content_type="text/html",
                charset="utf-8",
                status=404,
            )

        ws = session.get("ws")
        account = str(session.get("account", "")).strip() or "(unknown)"
        if not isinstance(ws, web.WebSocketResponse) or ws.closed:
            return web.Response(
                text="<h3>不存在的客户端</h3><p>cookie 存在，但客户端当前不在线。</p>",
                content_type="text/html",
                charset="utf-8",
                status=404,
            )

        ok, data = await self._call_client_rpc(
            ws,
            method=self.RPC_METHOD_CLIENT_FORCE_LOGOUT,
            params={"cookie": cookie},
            timeout_seconds=20,
        )
        self.logger.info(
            "view logout request: mail_type=%s account=%s cookie=%s ok=%s",
            "outlook",
            account,
            cookie,
            ok,
        )
        if not ok:
            message = escape(str(data.get("message", "client rpc failed")))
            return web.Response(
                text=(
                    "<h3>客户端退出失败</h3>"
                    f"<p>{message}</p>"
                    "<p><a href=\"/view/mail/clients\">返回客户端列表</a></p>"
                ),
                content_type="text/html",
                charset="utf-8",
                status=502,
            )

        message = escape(str(data.get("message", "logout requested")))
        return web.Response(
            text=(
                "<h3>已触发客户端退出</h3>"
                f"<p>account: <b>{escape(account)}</b></p>"
                f"<p>cookie: <code>{escape(cookie)}</code></p>"
                f"<p>{message}</p>"
                "<p><a href=\"/view/mail/clients\">返回客户端列表</a></p>"
            ),
            content_type="text/html",
            charset="utf-8",
        )

    def _try_resolve_server_to_client_rpc_response(self, ws: web.WebSocketResponse, text: str) -> bool:
        try:
            payload = json.loads(text)
        except json.JSONDecodeError:
            return False
        if not isinstance(payload, dict):
            return False
        if payload.get("jsonrpc") != "2.0":
            return False
        rpc_id = payload.get("id")
        if not isinstance(rpc_id, int):
            return False
        if "result" not in payload and "error" not in payload:
            return False

        pending = self._server_to_client_pending.get(rpc_id)
        if pending is None:
            return False
        pending_ws, fut = pending
        if pending_ws is not ws or fut.done():
            return False
        self.logger.info(
            "server<-client rpc response matched: rpc_id=%s has_error=%s",
            rpc_id,
            "error" in payload,
        )
        fut.set_result(payload)
        return True

    def _cleanup_server_to_client_pending_by_ws(self, ws: web.WebSocketResponse) -> None:
        rpc_ids = [rpc_id for rpc_id, (pending_ws, _) in self._server_to_client_pending.items() if pending_ws is ws]
        for rpc_id in rpc_ids:
            pending = self._server_to_client_pending.pop(rpc_id, None)
            if pending is None:
                continue
            _, fut = pending
            if not fut.done():
                fut.set_exception(RuntimeError("client websocket disconnected"))

    async def _call_client_rpc(
        self,
        ws: web.WebSocketResponse,
        method: str,
        params: dict[str, Any],
        timeout_seconds: int = 8,
    ) -> tuple[bool, dict[str, Any]]:
        if self._loop is None:
            self.logger.error("server->client rpc aborted: loop not ready, method=%s", method)
            return False, {"message": "server loop not ready"}
        self._server_to_client_rpc_id += 1
        rpc_id = self._server_to_client_rpc_id
        future = self._loop.create_future()
        self._server_to_client_pending[rpc_id] = (ws, future)
        rpc_start = time.perf_counter()
        request_payload = {
            "jsonrpc": "2.0",
            "id": rpc_id,
            "method": method,
            "params": params,
            "unixtime_ms": self._now_ms(),
        }
        if method == self.RPC_METHOD_LOCAL_FOLDER_LIST:
            self.logger.info("server->client rpc request: %s", request_payload)
        try:
            await ws.send_json(request_payload)
            payload = await asyncio.wait_for(future, timeout=timeout_seconds)
            if "error" in payload:
                error = payload.get("error") or {}
                self.logger.warning(
                    "server->client rpc error response: method=%s rpc_id=%s error=%s",
                    method,
                    rpc_id,
                    error,
                )
                return False, {"message": str(error.get("message", "client error"))}
            result = payload.get("result")
            if not isinstance(result, dict):
                self.logger.warning(
                    "server->client rpc invalid result type: method=%s rpc_id=%s result_type=%s",
                    method,
                    rpc_id,
                    type(result).__name__,
                )
                return False, {"message": "client result is not object"}
            elapsed_ms = int((time.perf_counter() - rpc_start) * 1000)
            if method == self.RPC_METHOD_LOCAL_FOLDER_LIST:
                folders = result.get("folders", [])
                folder_count = len(folders) if isinstance(folders, list) else 0
                self.logger.info(
                    "perf rpc folders.local.list(server->client): elapsed_ms=%s rpc_id=%s folder_count=%s",
                    elapsed_ms,
                    rpc_id,
                    folder_count,
                )
            self.logger.info(
                "server->client rpc success: method=%s rpc_id=%s keys=%s",
                method,
                rpc_id,
                list(result.keys()),
            )
            return True, result
        except asyncio.TimeoutError:
            self.logger.warning(
                "server->client rpc timeout: method=%s rpc_id=%s timeout=%ss",
                method,
                rpc_id,
                timeout_seconds,
            )
            return False, {"message": "client rpc timeout"}
        except Exception as exc:
            self.logger.exception(
                "server->client rpc exception: method=%s rpc_id=%s",
                method,
                rpc_id,
            )
            return False, {"message": str(exc)}
        finally:
            self._server_to_client_pending.pop(rpc_id, None)

    async def _dispatch_rpc_text(self, text: str, current_cookie: str) -> tuple[dict[str, Any] | None, str | None]:
        try:
            rpc = json.loads(text)
        except json.JSONDecodeError:
            return self._rpc_error(None, -32700, "Parse error", None), None

        if not isinstance(rpc, dict) or rpc.get("jsonrpc") != "2.0":
            rpc_id = rpc.get("id") if isinstance(rpc, dict) else None
            req_unixtime = self._normalize_unixtime_ms(rpc.get("unixtime_ms") if isinstance(rpc, dict) else None)
            return self._rpc_error(rpc_id, -32600, "Invalid Request", req_unixtime), None

        rpc_id = rpc.get("id")
        req_unixtime = self._normalize_unixtime_ms(rpc.get("unixtime_ms"))
        method = str(rpc.get("method", "")).strip()
        params = rpc.get("params", {})
        if not isinstance(params, dict):
            return self._rpc_error(rpc_id, -32602, "Invalid params", req_unixtime), None

        if method == self.RPC_METHOD_LOGIN:
            result = self._rpc_auth_login(params)
            if "cookie" in result:
                return self._rpc_result(rpc_id, result, req_unixtime), str(result["cookie"])
            return self._rpc_error(rpc_id, -32001, result.get("message", "login failed"), req_unixtime), None

        if method == self.RPC_METHOD_CONFIRM:
            cookie = str(params.get("cookie", "")).strip()
            if not cookie:
                return self._rpc_error(rpc_id, -32602, "cookie required", req_unixtime), None
            ok, data = await asyncio.to_thread(self._confirm_login_via_queue, cookie)
            if not ok:
                return self._rpc_error(rpc_id, -32003, data.get("message", "cookie invalid"), req_unixtime), None
            self._touch_session_last_query(cookie, method)
            return self._rpc_result(rpc_id, data, req_unixtime), cookie

        if method == self.RPC_METHOD_ACQUIRE:
            cookie = str(params.get("cookie", "")).strip()
            if not cookie:
                return self._rpc_error(rpc_id, -32602, "cookie required", req_unixtime), None
            acquire_start = time.perf_counter()
            ok, data = await asyncio.to_thread(self._acquire_token_and_fetch_folders, cookie)
            acquire_elapsed_ms = int((time.perf_counter() - acquire_start) * 1000)
            if ok:
                folders = data.get("folders", [])
                folder_count = len(folders) if isinstance(folders, list) else 0
                self.logger.info(
                    "perf rpc acquire(server): elapsed_ms=%s cookie=%s folder_count=%s token_cached=%s",
                    acquire_elapsed_ms,
                    cookie,
                    folder_count,
                    data.get("token_cached"),
                )
            else:
                self.logger.warning(
                    "perf rpc acquire(server) failed: elapsed_ms=%s cookie=%s message=%s",
                    acquire_elapsed_ms,
                    cookie,
                    data.get("message"),
                )
            if not ok:
                return self._rpc_error(rpc_id, -32004, data.get("message", "token acquire failed"), req_unixtime), None
            self._touch_session_last_query(cookie, method)
            return self._rpc_result(rpc_id, data, req_unixtime), cookie

        if method == self.RPC_METHOD_FOLDER_COUNT:
            cookie = str(params.get("cookie", "")).strip() or current_cookie
            folder_name = str(params.get("folder_name", "")).strip()
            if not cookie or not folder_name:
                return self._rpc_error(rpc_id, -32602, "cookie and folder_name required", req_unixtime), None

            current_count_raw = params.get("current_count", 0)
            try:
                current_count = int(current_count_raw)
            except (TypeError, ValueError):
                current_count = 0

            ok, data = await asyncio.to_thread(self._query_folder_count, cookie, folder_name, current_count)
            if not ok:
                return self._rpc_error(rpc_id, -32005, data.get("message", "folder count failed"), req_unixtime), None
            self._touch_session_last_query(cookie, method)
            return self._rpc_result(rpc_id, data, req_unixtime), cookie

        if method == self.RPC_METHOD_TITLE:
            cookie = str(params.get("cookie", "")).strip() or current_cookie
            folder_name = str(params.get("folder_name", "")).strip()
            if not cookie or not folder_name:
                return self._rpc_error(rpc_id, -32602, "cookie and folder_name required", req_unixtime), None

            known_max_uid = self._normalize_positive_int(params.get("known_max_uid"))
            incremental_count = self._normalize_non_negative_int(params.get("incremental_count"))
            ok, data = await asyncio.to_thread(
                self._query_folder_titles,
                cookie,
                folder_name,
                known_max_uid,
                incremental_count,
            )
            if not ok:
                return self._rpc_error(rpc_id, -32006, data.get("message", "folder titles failed"), req_unixtime), None
            self._touch_session_last_query(cookie, method)
            return self._rpc_result(rpc_id, data, req_unixtime), cookie

        if method == self.RPC_METHOD_TITLE_BASE64A:
            cookie = str(params.get("cookie", "")).strip() or current_cookie
            folder_name = str(params.get("folder_name", "")).strip()
            if not cookie or not folder_name:
                return self._rpc_error(rpc_id, -32602, "cookie and folder_name required", req_unixtime), None

            known_max_uid = self._normalize_positive_int(params.get("known_max_uid"))
            incremental_count = self._normalize_non_negative_int(params.get("incremental_count"))
            ok, data = await asyncio.to_thread(
                self._query_folder_titles_base64a,
                cookie,
                folder_name,
                known_max_uid,
                incremental_count,
            )
            if not ok:
                return self._rpc_error(rpc_id, -32009, data.get("message", "folder base64a failed"), req_unixtime), None
            self._touch_session_last_query(cookie, method)
            return self._rpc_result(rpc_id, data, req_unixtime), cookie

        if method == self.RPC_METHOD_FEISHU_NOTIFY:
            cookie = str(params.get("cookie", "")).strip() or current_cookie
            body = str(params.get("body", "")).strip()
            title_raw = params.get("title")
            tag_raw = params.get("tag")
            title = str(title_raw).strip() if title_raw is not None else None
            tag = str(tag_raw).strip() if tag_raw is not None else None

            if not cookie:
                return self._rpc_error(rpc_id, -32602, "cookie required", req_unixtime), None
            if not body:
                return self._rpc_error(rpc_id, -32602, "body required", req_unixtime), None
            with self._session_lock:
                session = self._sessions.get(cookie)
            if session is None:
                return self._rpc_error(rpc_id, -32007, "invalid cookie", req_unixtime), None
            try:
                results = await send_feishu_message(self.logger, v_body=body, v_title=title, tag=tag)
                success_count = sum(1 for ok in results.values() if ok)
                self._touch_session_last_query(cookie, method)
                return self._rpc_result(
                    rpc_id,
                    {
                        "success": True,
                        "results": results,
                        "success_count": success_count,
                    },
                    req_unixtime,
                ), cookie
            except Exception as exc:
                self.logger.exception("feishu notify failed")
                return self._rpc_error(rpc_id, -32008, str(exc), req_unixtime), None

        if method == self.RPC_METHOD_LOGOUT:
            cookie = str(params.get("cookie", "")).strip() or current_cookie
            if not cookie:
                return self._rpc_error(rpc_id, -32602, "cookie required", req_unixtime), None
            self._delete_cookie_session(cookie, reason="client_logout")
            return self._rpc_result(rpc_id, {"success": True, "message": "logout success"}, req_unixtime), ""

        return self._rpc_error(rpc_id, -32601, f"Method not found: {method}", req_unixtime), None

    def _rpc_auth_login(self, params: dict[str, Any]) -> dict[str, Any]:
        account = str(params.get("account", "")).strip()
        password = str(params.get("password", "")).strip()
        if not account or not password:
            return {"success": False, "message": "account/password required"}

        cookie = secrets.token_urlsafe(24)
        login_at = datetime.now(timezone.utc).isoformat()
        with self._session_lock:
            self._sessions[cookie] = {
                "cookie": cookie,
                "account": account,
                "login_at": login_at,
                "last_query_at": "",
                "last_query_unixtime_ms": 0,
                "last_query_method": "",
                "confirmed": False,
                "access_token": "",
                "token_acquired_at": "",
                "folders": [],
                "imap_client": None,
                "imap_lock": threading.Lock(),
            }
        return {"success": True, "cookie": cookie, "login_at": login_at}

    def _touch_session_last_query(self, cookie: str, method: str) -> None:
        now_dt = datetime.now(timezone.utc)
        with self._session_lock:
            session = self._sessions.get(cookie)
            if session is None:
                return
            session["last_query_at"] = now_dt.isoformat()
            session["last_query_unixtime_ms"] = int(now_dt.timestamp() * 1000)
            session["last_query_method"] = method

    def _confirm_login_via_queue(self, cookie: str) -> tuple[bool, dict[str, Any]]:
        with self._session_lock:
            session = self._sessions.get(cookie)
            if session is None:
                return False, {"message": "invalid cookie"}

        if not self._secure_queue_roundtrip({"type": "login_confirm", "cookie": cookie}):
            return False, {"message": "secure queue timeout during login confirm"}

        return True, {
            "success": True,
            "cookie": cookie,
            "enabled_methods": [
                self.RPC_METHOD_ACQUIRE,
                self.RPC_METHOD_FOLDER_COUNT,
                self.RPC_METHOD_TITLE,
                self.RPC_METHOD_TITLE_BASE64A,
                self.RPC_METHOD_FEISHU_NOTIFY,
                self.RPC_METHOD_LOGOUT,
            ],
        }

    def _acquire_token_and_fetch_folders(self, cookie: str) -> tuple[bool, dict[str, Any]]:
        acquire_start = time.perf_counter()
        with self._session_lock:
            session = self._sessions.get(cookie)
            if session is None:
                return False, {"message": "invalid cookie"}
            cached_token = str(session.get("access_token", "")).strip()
            imap_lock = session.get("imap_lock")
            if not hasattr(imap_lock, "acquire") or not hasattr(imap_lock, "release"):
                imap_lock = threading.Lock()
                session["imap_lock"] = imap_lock

        try:
            build_service_start = time.perf_counter()
            service = self._build_outlook_service()
            build_service_ms = int((time.perf_counter() - build_service_start) * 1000)

            token_from_cache = bool(cached_token)
            access_token = cached_token
            token_acquire_ms = 0
            token_source = "session_cache"
            if not access_token:
                token_start = time.perf_counter()
                if service.config.refresh_token:
                    token_source = "refresh_token"
                    access_token = service.acquire_access_token_by_refresh_token()
                else:
                    token_source = "device_flow_or_token_cache"
                    access_token = service.acquire_access_token()
                token_acquire_ms = int((time.perf_counter() - token_start) * 1000)

            fetch_folders_start = time.perf_counter()
            with imap_lock:
                folders, imap_reused = self._query_mailbox_folders_with_reconnect(
                    session=session,
                    service=service,
                    access_token=access_token,
                )
            fetch_folders_ms = int((time.perf_counter() - fetch_folders_start) * 1000)

            if not token_from_cache:
                ok = self._secure_queue_roundtrip(
                    {
                        "type": "token_update",
                        "cookie": cookie,
                        "access_token": access_token,
                        "folders": folders,
                    }
                )
                if not ok:
                    return False, {"message": "secure queue timeout during token update"}
            else:
                with self._session_lock:
                    session = self._sessions.get(cookie)
                    if session:
                        session["folders"] = folders

            total_ms = int((time.perf_counter() - acquire_start) * 1000)
            self.logger.info(
                "perf acquire breakdown: cookie=%s token_source=%s token_cached=%s imap_reused=%s build_service_ms=%s token_ms=%s fetch_folders_ms=%s total_ms=%s folder_count=%s",
                cookie,
                token_source,
                token_from_cache,
                imap_reused,
                build_service_ms,
                token_acquire_ms,
                fetch_folders_ms,
                total_ms,
                len(folders),
            )
            return True, {
                "success": True,
                "cookie": cookie,
                "token_cached": token_from_cache,
                "token_preview": self._mask_token(access_token),
                "folders": folders,
            }
        except Exception as exc:
            self.logger.exception("outlook token/folders handling failed")
            return False, {"message": str(exc)}

    def _build_outlook_service(self) -> imap_outlook_oauth2.OutlookMailService:
        args = SimpleNamespace(
            config=imap_outlook_oauth2.resolve_default_config_path(),
            profile=os.getenv("OUTLOOK_PROFILE", "outlook"),
            mailbox="INBOX",
        )
        config = imap_outlook_oauth2.build_runtime_config(args)
        return imap_outlook_oauth2.OutlookMailService(config, self.logger)

    def _query_folder_count(self, cookie: str, folder_name: str, current_count: int) -> tuple[bool, dict[str, Any]]:
        with self._session_lock:
            session = self._sessions.get(cookie)
            if session is None:
                return False, {"message": "invalid cookie"}
            access_token = str(session.get("access_token", "")).strip()
            imap_lock = session.get("imap_lock")
            if not hasattr(imap_lock, "acquire") or not hasattr(imap_lock, "release"):
                imap_lock = threading.Lock()
                session["imap_lock"] = imap_lock

        if not access_token:
            return False, {"message": "token not ready, call outlook.token.acquire first"}

        try:
            service = self._build_outlook_service()
            with imap_lock:
                folder_count = self._query_folder_count_with_reconnect(
                    session=session,
                    service=service,
                    access_token=access_token,
                    folder_name=folder_name,
                )
            return True, {
                "success": True,
                "folder_name": folder_name,
                "request_current_count": current_count,
                "folder_count": folder_count,
                "update_unixtime_ms": self._now_ms(),
            }
        except Exception as exc:
            self.logger.exception("query folder count failed: folder=%s", folder_name)
            return False, {"message": str(exc)}

    def _query_folder_titles(
        self,
        cookie: str,
        folder_name: str,
        known_max_uid: int | None = None,
        incremental_count: int | None = None,
    ) -> tuple[bool, dict[str, Any]]:
        with self._session_lock:
            session = self._sessions.get(cookie)
            if session is None:
                return False, {"message": "invalid cookie"}
            access_token = str(session.get("access_token", "")).strip()
            imap_lock = session.get("imap_lock")
            if not hasattr(imap_lock, "acquire") or not hasattr(imap_lock, "release"):
                imap_lock = threading.Lock()
                session["imap_lock"] = imap_lock

        if not access_token:
            return False, {"message": "token not ready, call outlook.token.acquire first"}

        try:
            t_start = time.perf_counter()
            service = self._build_outlook_service()
            with imap_lock:
                titles = self._query_folder_titles_with_reconnect(
                    session=session,
                    service=service,
                    access_token=access_token,
                    folder_name=folder_name,
                    known_max_uid=known_max_uid,
                    incremental_count=incremental_count,
                )
            elapsed_ms = int((time.perf_counter() - t_start) * 1000)
            self.logger.info(
                "perf query title: folder=%s known_max_uid=%s incremental_count=%s returned=%s elapsed_ms=%s",
                folder_name,
                known_max_uid,
                incremental_count,
                len(titles),
                elapsed_ms,
            )
            return True, {
                "success": True,
                "folder_name": folder_name,
                "titles": titles,
                "update_unixtime_ms": self._now_ms(),
            }
        except Exception as exc:
            self.logger.exception("query folder titles failed: folder=%s", folder_name)
            return False, {"message": str(exc)}

    def _query_folder_titles_base64a(
        self,
        cookie: str,
        folder_name: str,
        known_max_uid: int | None = None,
        incremental_count: int | None = None,
    ) -> tuple[bool, dict[str, Any]]:
        with self._session_lock:
            session = self._sessions.get(cookie)
            if session is None:
                return False, {"message": "invalid cookie"}
            access_token = str(session.get("access_token", "")).strip()
            imap_lock = session.get("imap_lock")
            if not hasattr(imap_lock, "acquire") or not hasattr(imap_lock, "release"):
                imap_lock = threading.Lock()
                session["imap_lock"] = imap_lock

        if not access_token:
            return False, {"message": "token not ready, call outlook.token.acquire first"}

        try:
            t_start = time.perf_counter()
            service = self._build_outlook_service()
            with imap_lock:
                titles = self._query_folder_titles_base64a_with_reconnect(
                    session=session,
                    service=service,
                    access_token=access_token,
                    folder_name=folder_name,
                    known_max_uid=known_max_uid,
                    incremental_count=incremental_count,
                )
            elapsed_ms = int((time.perf_counter() - t_start) * 1000)
            self.logger.info(
                "perf query title.base64a: folder=%s known_max_uid=%s incremental_count=%s returned=%s elapsed_ms=%s",
                folder_name,
                known_max_uid,
                incremental_count,
                len(titles),
                elapsed_ms,
            )
            return True, {
                "success": True,
                "folder_name": folder_name,
                "titles": titles,
                "update_unixtime_ms": self._now_ms(),
            }
        except Exception as exc:
            self.logger.exception("query folder titles base64a failed: folder=%s", folder_name)
            return False, {"message": str(exc)}

    def _fetch_mailbox_folders(
        self,
        service: imap_outlook_oauth2.OutlookMailService,
        imap,
    ) -> tuple[list[dict[str, Any]], int]:
        self.logger.info("query mailbox folders by token")
        t_list_start = time.perf_counter()
        raw_folders = service.list_mailboxes(imap)
        list_ms = int((time.perf_counter() - t_list_start) * 1000)
        # 统一输出结构：每个文件夹仅包含 name 和 flags。
        normalized: list[dict[str, Any]] = []
        for item in raw_folders:
            name = str(item.get("name", ""))
            flags = [str(x) for x in item.get("flags", [])]
            normalized.append({"name": name, "flags": flags})
        return normalized, list_ms

    def _query_mailbox_folders_with_reconnect(
        self,
        session: dict[str, Any],
        service: imap_outlook_oauth2.OutlookMailService,
        access_token: str,
    ) -> tuple[list[dict[str, Any]], bool]:
        for attempt in range(2):
            reused_connection = True
            connect_ms = 0
            auth_ms = 0
            imap = session.get("imap_client")
            if imap is None:
                reused_connection = False
                t_connect_start = time.perf_counter()
                imap = imaplib.IMAP4_SSL(service.config.host, service.config.port)
                t_connected = time.perf_counter()
                xoauth2 = service.build_xoauth2(service.config.email_addr, access_token)
                imap.authenticate("XOAUTH2", lambda _: xoauth2)
                t_authed = time.perf_counter()
                connect_ms = int((t_connected - t_connect_start) * 1000)
                auth_ms = int((t_authed - t_connected) * 1000)
                session["imap_client"] = imap
            try:
                folders, list_ms = self._fetch_mailbox_folders(service, imap)
                self.logger.info(
                    "perf acquire imap list: reused=%s connect_ms=%s auth_ms=%s list_ms=%s total_ms=%s folder_count=%s",
                    reused_connection,
                    connect_ms,
                    auth_ms,
                    list_ms,
                    connect_ms + auth_ms + list_ms,
                    len(folders),
                )
                return folders, reused_connection
            except (imaplib.IMAP4.abort, imaplib.IMAP4.error, OSError, EOFError) as exc:
                self.logger.warning(
                    "mailbox folders query hit imap connection issue, will retry=%s err=%s",
                    attempt == 0,
                    exc,
                )
                self._drop_session_imap(session)
                if attempt == 0:
                    continue
                raise
        raise RuntimeError("unreachable")

    def _query_folder_count_with_reconnect(
        self,
        session: dict[str, Any],
        service: imap_outlook_oauth2.OutlookMailService,
        access_token: str,
        folder_name: str,
    ) -> int:
        for attempt in range(2):
            imap = session.get("imap_client")
            if imap is None:
                imap = self._create_authenticated_imap(service, access_token)
                session["imap_client"] = imap
            try:
                return self._fetch_folder_count(service, imap, folder_name)
            except (imaplib.IMAP4.abort, imaplib.IMAP4.error, OSError, EOFError) as exc:
                self.logger.warning(
                    "folder count query hit imap connection issue, will retry=%s folder=%s err=%s",
                    attempt == 0,
                    folder_name,
                    exc,
                )
                self._drop_session_imap(session)
                if attempt == 0:
                    continue
                raise
        raise RuntimeError("unreachable")

    @staticmethod
    def _fetch_folder_count(service: imap_outlook_oauth2.OutlookMailService, imap, folder_name: str) -> int:
        mailbox_name, _ = service.resolve_mailbox_name(imap, folder_name)
        typ, _ = imap.select(mailbox_name)
        if typ != "OK":
            raise RuntimeError(f"select mailbox failed: {folder_name}")
        typ, data = imap.search(None, "ALL")
        if typ != "OK":
            raise RuntimeError(f"search mailbox failed: {folder_name}")
        if not data or not data[0]:
            return 0
        return len(data[0].split())

    def _fetch_folder_titles(
        self,
        service: imap_outlook_oauth2.OutlookMailService,
        imap,
        folder_name: str,
        known_max_uid: int | None = None,
        incremental_count: int | None = None,
    ) -> list[dict[str, Any]]:
        t_query_start = time.perf_counter()
        mailbox_name, _ = service.resolve_mailbox_name(imap, folder_name)
        typ, _ = imap.select(mailbox_name)
        if typ != "OK":
            raise RuntimeError(f"select mailbox failed: {folder_name}")

        uid_mode = bool(known_max_uid and known_max_uid > 0)
        t_search_start = time.perf_counter()
        if uid_mode:
            typ, data = imap.uid("SEARCH", None, f"UID {known_max_uid + 1}:*")
            if typ != "OK":
                raise RuntimeError(f"uid search mailbox failed: {folder_name}")
            uid_tokens = (data[0].split() if data and data[0] else [])
            if not uid_tokens:
                search_ms = int((time.perf_counter() - t_search_start) * 1000)
                total_ms = int((time.perf_counter() - t_query_start) * 1000)
                self.logger.info(
                    "perf fetch title: folder=%s uid_mode=%s known_max_uid=%s incremental_count=%s matched=0 planned_fetch=0 fetched=0 search_ms=%s fetch_ms=0 total_ms=%s",
                    folder_name,
                    uid_mode,
                    known_max_uid,
                    incremental_count,
                    search_ms,
                    total_ms,
                )
                return []
            matched_count = len(uid_tokens)
            if incremental_count is not None and incremental_count > 0 and len(uid_tokens) > incremental_count:
                uid_tokens = uid_tokens[-incremental_count:]
            msg_tokens = uid_tokens
        else:
            typ, data = imap.search(None, "ALL")
            if typ != "OK":
                raise RuntimeError(f"search mailbox failed: {folder_name}")
            if not data or not data[0]:
                search_ms = int((time.perf_counter() - t_search_start) * 1000)
                total_ms = int((time.perf_counter() - t_query_start) * 1000)
                self.logger.info(
                    "perf fetch title: folder=%s uid_mode=%s known_max_uid=%s incremental_count=%s matched=0 planned_fetch=0 fetched=0 search_ms=%s fetch_ms=0 total_ms=%s",
                    folder_name,
                    uid_mode,
                    known_max_uid,
                    incremental_count,
                    search_ms,
                    total_ms,
                )
                return []
            msg_tokens = data[0].split()
            matched_count = len(msg_tokens)
        search_ms = int((time.perf_counter() - t_search_start) * 1000)
        planned_fetch_count = len(msg_tokens)

        rows: list[dict[str, Any]] = []
        sender_missing_count = 0
        t_fetch_start = time.perf_counter()
        for msg_token in msg_tokens:
            if uid_mode:
                typ, msg_data = imap.uid("FETCH", msg_token, "(UID BODY.PEEK[HEADER.FIELDS (DATE SUBJECT FROM MESSAGE-ID)])")
            else:
                typ, msg_data = imap.fetch(msg_token, "(UID BODY.PEEK[HEADER.FIELDS (DATE SUBJECT FROM MESSAGE-ID)])")
            if typ != "OK" or not msg_data or not msg_data[0] or not msg_data[0][1]:
                continue

            msg = email.message_from_bytes(msg_data[0][1] or b"")
            title = service.decode_mime_words(msg.get("Subject", "")) or "(无主题)"
            raw_from = str(msg.get("From", "") or "").strip()
            decoded_from = service.decode_mime_words(raw_from) if raw_from else ""
            _, sender_email = parseaddr(decoded_from or raw_from)
            sender = decoded_from or sender_email or raw_from
            sender = " ".join(str(sender).split())
            if not sender:
                sender_missing_count += 1
            date_raw = str(msg.get("Date", "")).strip()
            received_unixtime_ms = 0
            received_at = ""
            if date_raw:
                try:
                    dt = parsedate_to_datetime(date_raw)
                    if dt is not None:
                        if dt.tzinfo is None:
                            dt = dt.replace(tzinfo=timezone.utc)
                        dt = dt.astimezone(timezone.utc)
                        received_unixtime_ms = int(dt.timestamp() * 1000)
                        received_at = dt.isoformat()
                except Exception:
                    pass

            uid = ""
            fetch_meta = msg_data[0][0]
            if isinstance(fetch_meta, bytes):
                meta_text = fetch_meta.decode("ascii", errors="ignore")
                uid_match = re.search(r"\bUID\s+(\d+)\b", meta_text)
                if uid_match:
                    uid = uid_match.group(1)

            message_id = str(msg.get("Message-ID", "") or "").strip()
            rows.append(
                {
                    "mail_id": msg_token.decode("ascii", errors="ignore"),
                    "uid": uid,
                    "message_id": message_id,
                    "title": title,
                    "sender": str(sender or ""),
                    "received_at": received_at,
                    "received_unixtime_ms": received_unixtime_ms,
                }
            )
        self.logger.info(
            "folder titles fetched: folder=%s total=%s sender_missing=%s",
            folder_name,
            len(rows),
            sender_missing_count,
        )
        if sender_missing_count > 0:
            self.logger.warning(
                "some mails have empty From header: folder=%s missing=%s",
                folder_name,
                sender_missing_count,
            )
        fetch_ms = int((time.perf_counter() - t_fetch_start) * 1000)
        total_ms = int((time.perf_counter() - t_query_start) * 1000)
        self.logger.info(
            "perf fetch title: folder=%s uid_mode=%s known_max_uid=%s incremental_count=%s matched=%s planned_fetch=%s fetched=%s search_ms=%s fetch_ms=%s total_ms=%s",
            folder_name,
            uid_mode,
            known_max_uid,
            incremental_count,
            matched_count,
            planned_fetch_count,
            len(rows),
            search_ms,
            fetch_ms,
            total_ms,
        )
        return rows

    def _fetch_folder_titles_base64a(
        self,
        service: imap_outlook_oauth2.OutlookMailService,
        imap,
        folder_name: str,
        known_max_uid: int | None = None,
        incremental_count: int | None = None,
    ) -> list[dict[str, Any]]:
        t_query_start = time.perf_counter()
        mailbox_name, _ = service.resolve_mailbox_name(imap, folder_name)
        typ, _ = imap.select(mailbox_name)
        if typ != "OK":
            raise RuntimeError(f"select mailbox failed: {folder_name}")

        uid_mode = bool(known_max_uid and known_max_uid > 0)
        t_search_start = time.perf_counter()
        if uid_mode:
            typ, data = imap.uid("SEARCH", None, f"UID {known_max_uid + 1}:*")
            if typ != "OK":
                raise RuntimeError(f"uid search mailbox failed: {folder_name}")
            uid_tokens = (data[0].split() if data and data[0] else [])
            if not uid_tokens:
                search_ms = int((time.perf_counter() - t_search_start) * 1000)
                total_ms = int((time.perf_counter() - t_query_start) * 1000)
                self.logger.info(
                    "perf fetch title.base64a: folder=%s uid_mode=%s known_max_uid=%s incremental_count=%s matched=0 planned_fetch=0 fetched=0 search_ms=%s fetch_ms=0 total_ms=%s",
                    folder_name,
                    uid_mode,
                    known_max_uid,
                    incremental_count,
                    search_ms,
                    total_ms,
                )
                return []
            matched_count = len(uid_tokens)
            if incremental_count is not None and incremental_count > 0 and len(uid_tokens) > incremental_count:
                uid_tokens = uid_tokens[-incremental_count:]
            msg_tokens = uid_tokens
        else:
            typ, data = imap.search(None, "ALL")
            if typ != "OK":
                raise RuntimeError(f"search mailbox failed: {folder_name}")
            if not data or not data[0]:
                search_ms = int((time.perf_counter() - t_search_start) * 1000)
                total_ms = int((time.perf_counter() - t_query_start) * 1000)
                self.logger.info(
                    "perf fetch title.base64a: folder=%s uid_mode=%s known_max_uid=%s incremental_count=%s matched=0 planned_fetch=0 fetched=0 search_ms=%s fetch_ms=0 total_ms=%s",
                    folder_name,
                    uid_mode,
                    known_max_uid,
                    incremental_count,
                    search_ms,
                    total_ms,
                )
                return []
            msg_tokens = data[0].split()
            matched_count = len(msg_tokens)
        search_ms = int((time.perf_counter() - t_search_start) * 1000)
        planned_fetch_count = len(msg_tokens)

        rows: list[dict[str, Any]] = []
        sender_missing_count = 0
        t_fetch_start = time.perf_counter()
        for msg_token in msg_tokens:
            if uid_mode:
                typ, msg_data = imap.uid("FETCH", msg_token, "(UID BODY.PEEK[])")
            else:
                typ, msg_data = imap.fetch(msg_token, "(UID BODY.PEEK[])")
            if typ != "OK" or not msg_data or not msg_data[0] or not msg_data[0][1]:
                continue
            raw_message = msg_data[0][1] or b""
            msg = email.message_from_bytes(raw_message)
            title = service.decode_mime_words(msg.get("Subject", "")) or "(无主题)"
            raw_from = str(msg.get("From", "") or "").strip()
            decoded_from = service.decode_mime_words(raw_from) if raw_from else ""
            _, sender_email = parseaddr(decoded_from or raw_from)
            sender = decoded_from or sender_email or raw_from
            sender = " ".join(str(sender).split())
            if not sender:
                sender_missing_count += 1
            date_raw = str(msg.get("Date", "")).strip()
            received_unixtime_ms = 0
            received_at = ""
            if date_raw:
                try:
                    dt = parsedate_to_datetime(date_raw)
                    if dt is not None:
                        if dt.tzinfo is None:
                            dt = dt.replace(tzinfo=timezone.utc)
                        dt = dt.astimezone(timezone.utc)
                        received_unixtime_ms = int(dt.timestamp() * 1000)
                        received_at = dt.isoformat()
                except Exception:
                    pass

            uid = ""
            fetch_meta = msg_data[0][0]
            if isinstance(fetch_meta, bytes):
                meta_text = fetch_meta.decode("ascii", errors="ignore")
                uid_match = re.search(r"\bUID\s+(\d+)\b", meta_text)
                if uid_match:
                    uid = uid_match.group(1)

            message_id = str(msg.get("Message-ID", "") or "").strip()
            rows.append(
                {
                    "mail_id": msg_token.decode("ascii", errors="ignore"),
                    "uid": uid,
                    "message_id": message_id,
                    "title": title,
                    "sender": str(sender or ""),
                    "received_at": received_at,
                    "received_unixtime_ms": received_unixtime_ms,
                    "Base64A": base64.urlsafe_b64encode(raw_message).replace(b"=", b"").decode("utf-8"),
                }
            )
        self.logger.info(
            "folder base64a fetched: folder=%s total=%s sender_missing=%s",
            folder_name,
            len(rows),
            sender_missing_count,
        )
        if sender_missing_count > 0:
            self.logger.warning(
                "some mails have empty From header(base64a): folder=%s missing=%s",
                folder_name,
                sender_missing_count,
            )
        fetch_ms = int((time.perf_counter() - t_fetch_start) * 1000)
        total_ms = int((time.perf_counter() - t_query_start) * 1000)
        self.logger.info(
            "perf fetch title.base64a: folder=%s uid_mode=%s known_max_uid=%s incremental_count=%s matched=%s planned_fetch=%s fetched=%s search_ms=%s fetch_ms=%s total_ms=%s",
            folder_name,
            uid_mode,
            known_max_uid,
            incremental_count,
            matched_count,
            planned_fetch_count,
            len(rows),
            search_ms,
            fetch_ms,
            total_ms,
        )
        return rows

    def _query_folder_titles_with_reconnect(
        self,
        session: dict[str, Any],
        service: imap_outlook_oauth2.OutlookMailService,
        access_token: str,
        folder_name: str,
        known_max_uid: int | None = None,
        incremental_count: int | None = None,
    ) -> list[dict[str, Any]]:
        for attempt in range(2):
            imap = session.get("imap_client")
            if imap is None:
                imap = self._create_authenticated_imap(service, access_token)
                session["imap_client"] = imap
            try:
                return self._fetch_folder_titles(
                    service,
                    imap,
                    folder_name,
                    known_max_uid=known_max_uid,
                    incremental_count=incremental_count,
                )
            except (imaplib.IMAP4.abort, imaplib.IMAP4.error, OSError, EOFError) as exc:
                self.logger.warning(
                    "folder titles query hit imap connection issue, will retry=%s folder=%s err=%s",
                    attempt == 0,
                    folder_name,
                    exc,
                )
                self._drop_session_imap(session)
                if attempt == 0:
                    continue
                raise
        raise RuntimeError("unreachable")

    def _query_folder_titles_base64a_with_reconnect(
        self,
        session: dict[str, Any],
        service: imap_outlook_oauth2.OutlookMailService,
        access_token: str,
        folder_name: str,
        known_max_uid: int | None = None,
        incremental_count: int | None = None,
    ) -> list[dict[str, Any]]:
        for attempt in range(2):
            imap = session.get("imap_client")
            if imap is None:
                imap = self._create_authenticated_imap(service, access_token)
                session["imap_client"] = imap
            try:
                return self._fetch_folder_titles_base64a(
                    service,
                    imap,
                    folder_name,
                    known_max_uid=known_max_uid,
                    incremental_count=incremental_count,
                )
            except (imaplib.IMAP4.abort, imaplib.IMAP4.error, OSError, EOFError) as exc:
                self.logger.warning(
                    "folder base64a query hit imap connection issue, will retry=%s folder=%s err=%s",
                    attempt == 0,
                    folder_name,
                    exc,
                )
                self._drop_session_imap(session)
                if attempt == 0:
                    continue
                raise
        raise RuntimeError("unreachable")

    @staticmethod
    def _create_authenticated_imap(service: imap_outlook_oauth2.OutlookMailService, access_token: str):
        imap = imaplib.IMAP4_SSL(service.config.host, service.config.port)
        xoauth2 = service.build_xoauth2(service.config.email_addr, access_token)
        imap.authenticate("XOAUTH2", lambda _: xoauth2)
        return imap

    def _drop_session_imap(self, session: dict[str, Any]) -> None:
        imap = session.pop("imap_client", None)
        if imap is not None:
            try:
                imap.logout()
            except Exception:
                pass

    def _delete_cookie_session(self, cookie: str, reason: str) -> None:
        with self._session_lock:
            removed = self._sessions.pop(cookie, None)
        if removed is not None:
            removed.pop("ws", None)
            lock = removed.get("imap_lock")
            if hasattr(lock, "acquire") and hasattr(lock, "release"):
                if lock.acquire(timeout=2):
                    try:
                        self._drop_session_imap(removed)
                    finally:
                        lock.release()
            else:
                self._drop_session_imap(removed)
            self.logger.info("cookie deleted: reason=%s", reason)

    def _secure_queue_roundtrip(self, payload: dict[str, Any]) -> bool:
        ack = threading.Event()
        event = dict(payload)
        event["ack"] = ack
        self._secure_queue.put(event)
        ok = ack.wait(timeout=self.QUEUE_ACK_TIMEOUT_SECONDS)
        if not ok:
            self.logger.warning("secure queue ack timeout: type=%s", payload.get("type"))
        return ok

    def _mark_client_connected(self) -> None:
        with self._clients_lock:
            self._active_clients += 1

    def _mark_client_disconnected(self) -> None:
        with self._clients_lock:
            self._active_clients = max(0, self._active_clients - 1)

    @staticmethod
    def _mask_token(token: str) -> str:
        if not token:
            return ""
        if len(token) <= 10:
            return "*" * len(token)
        return f"{token[:6]}***{token[-4:]}"

    @staticmethod
    def _rpc_result(rpc_id: Any, result: dict[str, Any], request_unixtime_ms: int | None) -> dict[str, Any]:
        return {
            "jsonrpc": "2.0",
            "id": rpc_id,
            "result": result,
            "request_unixtime_ms": request_unixtime_ms,
            "response_unixtime_ms": InternalWSServer._now_ms(),
        }

    @staticmethod
    def _rpc_error(rpc_id: Any, code: int, message: str, request_unixtime_ms: int | None) -> dict[str, Any]:
        return {
            "jsonrpc": "2.0",
            "id": rpc_id,
            "error": {"code": code, "message": message},
            "request_unixtime_ms": request_unixtime_ms,
            "response_unixtime_ms": InternalWSServer._now_ms(),
        }

    @staticmethod
    def _rpc_notification(method: str, params: dict[str, Any]) -> dict[str, Any]:
        return {
            "jsonrpc": "2.0",
            "method": method,
            "params": params,
            "request_unixtime_ms": None,
            "response_unixtime_ms": InternalWSServer._now_ms(),
        }

    @staticmethod
    def _normalize_unixtime_ms(value: Any) -> int | None:
        if value is None:
            return None
        try:
            normalized = int(value)
            if normalized < 0:
                return None
            return normalized
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _normalize_positive_int(value: Any) -> int | None:
        if value is None:
            return None
        try:
            normalized = int(value)
            if normalized <= 0:
                return None
            return normalized
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _normalize_non_negative_int(value: Any) -> int | None:
        if value is None:
            return None
        try:
            normalized = int(value)
            if normalized < 0:
                return None
            return normalized
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _now_ms() -> int:
        return int(time.time() * 1000)
