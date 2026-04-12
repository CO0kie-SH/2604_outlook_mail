import asyncio
import imaplib
import json
import os
import secrets
import threading
import time
from datetime import datetime, timezone
from queue import Empty, Queue
from types import SimpleNamespace
from typing import Any

from aiohttp import web

import imap_outlook_oauth2


class InternalWSServer:
    """邮件模块 WebSocket 服务端，使用 JSON-RPC 2.0 进行内部通信。"""

    IDLE_CHECK_INTERVAL_SECONDS = 60
    IDLE_ZERO_LIMIT = 2

    def __init__(self, host: str, port: int, logger):
        self.host = host
        self.port = port
        self.logger = logger

        self.ready_event = threading.Event()
        self.shutdown_requested_event = threading.Event()
        self._stop_event = threading.Event()

        self._thread: threading.Thread | None = None
        self._consumer_thread: threading.Thread | None = None
        self._idle_checker_thread: threading.Thread | None = None

        self._loop: asyncio.AbstractEventLoop | None = None
        self._runner: web.AppRunner | None = None

        # 安全队列：统一串行处理登录确认和 token 缓存等敏感会话操作。
        self._secure_queue: Queue[dict[str, Any]] = Queue()
        self._sessions: dict[str, dict[str, Any]] = {}
        self._session_lock = threading.Lock()

        self._clients_lock = threading.Lock()
        self._active_clients = 0

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return

        self._stop_event.clear()
        self._consumer_thread = threading.Thread(target=self._consume_secure_queue, name="ws-secure-queue", daemon=True)
        self._consumer_thread.start()

        # 每 60 秒检查一次客户端连接数，连续两次为 0 则触发退出。
        self._idle_checker_thread = threading.Thread(target=self._idle_exit_checker, name="ws-idle-checker", daemon=True)
        self._idle_checker_thread.start()

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
        app.add_routes([web.get("/ws/mail", self._handle_mail_ws)])
        self._runner = web.AppRunner(app)
        await self._runner.setup()
        site = web.TCPSite(self._runner, self.host, self.port)
        await site.start()

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
                            session["inbox_total"] = int(event.get("inbox_total", 0))
                else:
                    self.logger.warning("unknown secure queue event: %s", event_type)
            finally:
                if isinstance(ack_event, threading.Event):
                    ack_event.set()

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
                            "auth.login",
                            "auth.confirm",
                            "outlook.token.acquire",
                            "auth.logout",
                        ],
                    },
                )
            )

            async for msg in ws:
                if msg.type != web.WSMsgType.TEXT:
                    if msg.type in {web.WSMsgType.CLOSE, web.WSMsgType.CLOSING, web.WSMsgType.CLOSED}:
                        break
                    continue

                response, new_cookie = await self._dispatch_rpc_text(msg.data, current_cookie)
                if new_cookie is not None:
                    current_cookie = new_cookie
                if response is not None:
                    await ws.send_json(response)

        finally:
            if current_cookie:
                self._delete_cookie_session(current_cookie, reason="ws_disconnected")
            self._mark_client_disconnected()
            self.logger.info("mail websocket client disconnected")

        return ws

    async def _dispatch_rpc_text(self, text: str, current_cookie: str) -> tuple[dict[str, Any] | None, str | None]:
        try:
            rpc = json.loads(text)
        except json.JSONDecodeError:
            return self._rpc_error(None, -32700, "Parse error"), None

        if not isinstance(rpc, dict) or rpc.get("jsonrpc") != "2.0":
            return self._rpc_error(rpc.get("id") if isinstance(rpc, dict) else None, -32600, "Invalid Request"), None

        rpc_id = rpc.get("id")
        method = str(rpc.get("method", "")).strip()
        params = rpc.get("params", {})
        if not isinstance(params, dict):
            return self._rpc_error(rpc_id, -32602, "Invalid params"), None

        if method == "auth.login":
            result = self._rpc_auth_login(params)
            if "cookie" in result:
                return self._rpc_result(rpc_id, result), str(result["cookie"])
            return self._rpc_error(rpc_id, -32001, result.get("message", "login failed")), None

        if method == "auth.confirm":
            cookie = str(params.get("cookie", "")).strip()
            if not cookie:
                return self._rpc_error(rpc_id, -32602, "cookie required"), None
            ok, data = await asyncio.to_thread(self._confirm_login_via_queue, cookie)
            if not ok:
                return self._rpc_error(rpc_id, -32003, data.get("message", "cookie invalid")), None
            return self._rpc_result(rpc_id, data), cookie

        if method == "outlook.token.acquire":
            cookie = str(params.get("cookie", "")).strip()
            if not cookie:
                return self._rpc_error(rpc_id, -32602, "cookie required"), None
            ok, data = await asyncio.to_thread(self._acquire_token_and_query_inbox, cookie)
            if not ok:
                return self._rpc_error(rpc_id, -32004, data.get("message", "token acquire failed")), None
            return self._rpc_result(rpc_id, data), cookie

        if method == "auth.logout":
            cookie = str(params.get("cookie", "")).strip() or current_cookie
            if not cookie:
                return self._rpc_error(rpc_id, -32602, "cookie required"), None
            self._delete_cookie_session(cookie, reason="client_logout")
            return self._rpc_result(rpc_id, {"success": True, "message": "logout success"}), ""

        return self._rpc_error(rpc_id, -32601, f"Method not found: {method}"), None

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
                "confirmed": False,
                "access_token": "",
                "token_acquired_at": "",
                "inbox_total": None,
            }
        return {"success": True, "cookie": cookie, "login_at": login_at}

    def _confirm_login_via_queue(self, cookie: str) -> tuple[bool, dict[str, Any]]:
        with self._session_lock:
            session = self._sessions.get(cookie)
            if session is None:
                return False, {"message": "invalid cookie"}

        ack = threading.Event()
        self._secure_queue.put({"type": "login_confirm", "cookie": cookie, "ack": ack})
        ack.wait(timeout=3)

        return True, {
            "success": True,
            "cookie": cookie,
            "enabled_methods": [
                "outlook.token.acquire",
                "auth.logout",
            ],
        }

    def _acquire_token_and_query_inbox(self, cookie: str) -> tuple[bool, dict[str, Any]]:
        with self._session_lock:
            session = self._sessions.get(cookie)
            if session is None:
                return False, {"message": "invalid cookie"}
            cached_token = str(session.get("access_token", "")).strip()

        try:
            service = self._build_outlook_service()

            token_from_cache = bool(cached_token)
            access_token = cached_token
            if not access_token:
                access_token = (
                    service.acquire_access_token_by_refresh_token()
                    if service.config.refresh_token
                    else service.acquire_access_token()
                )

            inbox_total = self._fetch_inbox_total(service, access_token)

            if not token_from_cache:
                ack = threading.Event()
                self._secure_queue.put(
                    {
                        "type": "token_update",
                        "cookie": cookie,
                        "access_token": access_token,
                        "inbox_total": inbox_total,
                        "ack": ack,
                    }
                )
                ack.wait(timeout=3)
            else:
                with self._session_lock:
                    session = self._sessions.get(cookie)
                    if session:
                        session["inbox_total"] = inbox_total

            return True, {
                "success": True,
                "cookie": cookie,
                "token_cached": token_from_cache,
                "token_preview": self._mask_token(access_token),
                "mailbox": "INBOX",
                "inbox_total": inbox_total,
            }
        except Exception as exc:
            self.logger.exception("outlook token/count handling failed")
            return False, {"message": str(exc)}

    def _build_outlook_service(self) -> imap_outlook_oauth2.OutlookMailService:
        args = SimpleNamespace(
            config=imap_outlook_oauth2.resolve_default_config_path(),
            profile=os.getenv("OUTLOOK_PROFILE", "outlook"),
            mailbox="INBOX",
        )
        config = imap_outlook_oauth2.build_runtime_config(args)
        return imap_outlook_oauth2.OutlookMailService(config, self.logger)

    def _fetch_inbox_total(self, service: imap_outlook_oauth2.OutlookMailService, access_token: str) -> int:
        self.logger.info("query inbox total by token")
        imap = imaplib.IMAP4_SSL(service.config.host, service.config.port)
        try:
            xoauth2 = service.build_xoauth2(service.config.email_addr, access_token)
            imap.authenticate("XOAUTH2", lambda _: xoauth2)
            typ, _ = imap.select("INBOX")
            if typ != "OK":
                raise RuntimeError("select INBOX failed")
            typ, data = imap.search(None, "ALL")
            if typ != "OK":
                raise RuntimeError("search INBOX failed")
            if not data or not data[0]:
                return 0
            return len(data[0].split())
        finally:
            try:
                imap.logout()
            except Exception:
                pass

    def _delete_cookie_session(self, cookie: str, reason: str) -> None:
        with self._session_lock:
            removed = self._sessions.pop(cookie, None)
        if removed is not None:
            self.logger.info("cookie deleted: reason=%s", reason)

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
    def _rpc_result(rpc_id: Any, result: dict[str, Any]) -> dict[str, Any]:
        return {"jsonrpc": "2.0", "id": rpc_id, "result": result}

    @staticmethod
    def _rpc_error(rpc_id: Any, code: int, message: str) -> dict[str, Any]:
        return {"jsonrpc": "2.0", "id": rpc_id, "error": {"code": code, "message": message}}

    @staticmethod
    def _rpc_notification(method: str, params: dict[str, Any]) -> dict[str, Any]:
        return {"jsonrpc": "2.0", "method": method, "params": params}
