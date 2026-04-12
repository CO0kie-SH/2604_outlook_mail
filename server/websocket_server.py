import asyncio
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
from queue import Empty, Queue
from types import SimpleNamespace
from typing import Any

from aiohttp import web

from feishu_notifier import send_feishu_message
import imap_outlook_oauth2
from server.rpc_docs import InternalWSDocPages


class InternalWSServer:
    """邮件模块 WebSocket 服务端，使用 JSON-RPC 2.0 进行内部通信。"""

    IDLE_CHECK_INTERVAL_SECONDS = 30
    IDLE_ZERO_LIMIT = 2
    QUEUE_ACK_TIMEOUT_SECONDS = 3

    RPC_METHOD_LOGIN = "auth.login"
    RPC_METHOD_CONFIRM = "auth.confirm"
    RPC_METHOD_ACQUIRE = "outlook.token.acquire"
    RPC_METHOD_FOLDER_COUNT = "mail.folder.count"
    RPC_METHOD_TITLE = "title"
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

        self._loop: asyncio.AbstractEventLoop | None = None
        self._runner: web.AppRunner | None = None

        # 安全队列：串行处理登录确认和 token 缓存等敏感会话动作。
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

        # 每 30 秒检查一次客户端连接数，连续两次为 0 则退出。
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
        app.add_routes(
            [
                web.get("/ws/mail", self._handle_mail_ws),
                web.get("/doc/mail", self._doc_pages.handle_doc_mail),
                web.get("/doc/feishu", self._doc_pages.handle_doc_feishu),
            ]
        )
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
                            session["folders"] = event.get("folders", [])
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
                            self.RPC_METHOD_LOGIN,
                            self.RPC_METHOD_CONFIRM,
                            self.RPC_METHOD_ACQUIRE,
                            self.RPC_METHOD_FOLDER_COUNT,
                            self.RPC_METHOD_TITLE,
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
            return self._rpc_result(rpc_id, data, req_unixtime), cookie

        if method == self.RPC_METHOD_ACQUIRE:
            cookie = str(params.get("cookie", "")).strip()
            if not cookie:
                return self._rpc_error(rpc_id, -32602, "cookie required", req_unixtime), None
            ok, data = await asyncio.to_thread(self._acquire_token_and_fetch_folders, cookie)
            if not ok:
                return self._rpc_error(rpc_id, -32004, data.get("message", "token acquire failed"), req_unixtime), None
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
            return self._rpc_result(rpc_id, data, req_unixtime), cookie

        if method == self.RPC_METHOD_TITLE:
            cookie = str(params.get("cookie", "")).strip() or current_cookie
            folder_name = str(params.get("folder_name", "")).strip()
            if not cookie or not folder_name:
                return self._rpc_error(rpc_id, -32602, "cookie and folder_name required", req_unixtime), None

            ok, data = await asyncio.to_thread(self._query_folder_titles, cookie, folder_name)
            if not ok:
                return self._rpc_error(rpc_id, -32006, data.get("message", "folder titles failed"), req_unixtime), None
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
                "confirmed": False,
                "access_token": "",
                "token_acquired_at": "",
                "folders": [],
            }
        return {"success": True, "cookie": cookie, "login_at": login_at}

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
                self.RPC_METHOD_FEISHU_NOTIFY,
                self.RPC_METHOD_LOGOUT,
            ],
        }

    def _acquire_token_and_fetch_folders(self, cookie: str) -> tuple[bool, dict[str, Any]]:
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

            folders = self._fetch_mailbox_folders(service, access_token)

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

        if not access_token:
            return False, {"message": "token not ready, call outlook.token.acquire first"}

        try:
            service = self._build_outlook_service()
            folder_count = self._fetch_folder_count(service, access_token, folder_name)
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

    def _query_folder_titles(self, cookie: str, folder_name: str) -> tuple[bool, dict[str, Any]]:
        with self._session_lock:
            session = self._sessions.get(cookie)
            if session is None:
                return False, {"message": "invalid cookie"}
            access_token = str(session.get("access_token", "")).strip()

        if not access_token:
            return False, {"message": "token not ready, call outlook.token.acquire first"}

        try:
            service = self._build_outlook_service()
            titles = self._fetch_folder_titles(service, access_token, folder_name)
            return True, {
                "success": True,
                "folder_name": folder_name,
                "titles": titles,
                "update_unixtime_ms": self._now_ms(),
            }
        except Exception as exc:
            self.logger.exception("query folder titles failed: folder=%s", folder_name)
            return False, {"message": str(exc)}

    def _fetch_mailbox_folders(self, service: imap_outlook_oauth2.OutlookMailService, access_token: str) -> list[dict[str, Any]]:
        self.logger.info("query mailbox folders by token")
        imap = imaplib.IMAP4_SSL(service.config.host, service.config.port)
        try:
            xoauth2 = service.build_xoauth2(service.config.email_addr, access_token)
            imap.authenticate("XOAUTH2", lambda _: xoauth2)
            raw_folders = service.list_mailboxes(imap)
            # 统一输出结构：每个文件夹仅包含 name 和 flags。
            normalized: list[dict[str, Any]] = []
            for item in raw_folders:
                name = str(item.get("name", ""))
                flags = [str(x) for x in item.get("flags", [])]
                normalized.append({"name": name, "flags": flags})
            return normalized
        finally:
            try:
                imap.logout()
            except Exception:
                pass

    def _fetch_folder_count(self, service: imap_outlook_oauth2.OutlookMailService, access_token: str, folder_name: str) -> int:
        imap = imaplib.IMAP4_SSL(service.config.host, service.config.port)
        try:
            xoauth2 = service.build_xoauth2(service.config.email_addr, access_token)
            imap.authenticate("XOAUTH2", lambda _: xoauth2)
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
        finally:
            try:
                imap.logout()
            except Exception:
                pass

    def _fetch_folder_titles(
        self,
        service: imap_outlook_oauth2.OutlookMailService,
        access_token: str,
        folder_name: str,
    ) -> list[dict[str, Any]]:
        imap = imaplib.IMAP4_SSL(service.config.host, service.config.port)
        try:
            xoauth2 = service.build_xoauth2(service.config.email_addr, access_token)
            imap.authenticate("XOAUTH2", lambda _: xoauth2)
            mailbox_name, _ = service.resolve_mailbox_name(imap, folder_name)
            typ, _ = imap.select(mailbox_name)
            if typ != "OK":
                raise RuntimeError(f"select mailbox failed: {folder_name}")

            typ, data = imap.search(None, "ALL")
            if typ != "OK":
                raise RuntimeError(f"search mailbox failed: {folder_name}")
            if not data or not data[0]:
                return []

            rows: list[dict[str, Any]] = []
            sender_missing_count = 0
            for msg_id in data[0].split():
                typ, msg_data = imap.fetch(msg_id, "(UID BODY.PEEK[HEADER.FIELDS (DATE SUBJECT FROM MESSAGE-ID)])")
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
                        "mail_id": msg_id.decode("ascii", errors="ignore"),
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
            return rows
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
    def _now_ms() -> int:
        return int(time.time() * 1000)
