import asyncio
import csv
from contextlib import suppress
import json
import re
import stat
import threading
import time
from pathlib import Path
from typing import Any

import aiohttp


class InternalWSClient:
    """Mail module WebSocket client using JSON-RPC 2.0 flow."""
    RPC_METHOD_LOCAL_FOLDER_LIST = "mail.folders.local.list"
    RPC_METHOD_LOCAL_TITLE_LIST = "mail.titles.local.list"
    RPC_METHOD_CLIENT_FORCE_LOGOUT = "mail.client.force.logout"
    RPC_TIMEOUT_SECONDS = 20
    POST_FLOW_FOLDER_PULL_INTERVAL_SECONDS = 8
    POST_FLOW_FOLDER_PULL_TIMES = 5
    FOLDER_CSV_FIELDS = [
        "unixtime_ms",
        "name",
        "flags",
        "mode",
        "current_count",
        "online_count",
        "current_unixtime_ms",
        "update_unixtime_ms",
    ]
    TITLE_CSV_FIELDS = [
        "mail_id",
        "uid",
        "message_id",
        "sender",
        "title",
        "received_at",
        "received_unixtime_ms",
        "unixtime_ms",
        "Base64A",
    ]
    FEISHU_TITLE_LIMIT = 20
    FEISHU_PER_FOLDER_LIMIT = 5
    FEISHU_BODY_MAX_CHARS = 3500

    def __init__(
        self,
        server_host: str,
        server_port: int,
        account: str,
        password: str,
        logger,
    ):
        self.server_host = server_host
        self.server_port = server_port
        self.account = account
        self.password = password
        self.logger = logger

        self._thread: threading.Thread | None = None
        self._loop: asyncio.AbstractEventLoop | None = None
        self._stop_event = threading.Event()
        self._force_logout_requested = False
        self._logged_out = False
        self._ensure_csv_field_limit()

    @property
    def ws_url(self) -> str:
        return f"ws://{self.server_host}:{self.server_port}/ws/mail"

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run_client_thread, name="ws-mail-client", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        if self._loop and self._loop.is_running():
            self._loop.call_soon_threadsafe(self._loop.stop)
        if self._thread:
            self._thread.join(timeout=5)

    def _run_client_thread(self) -> None:
        self._loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self._loop)
        task = self._loop.create_task(self._run_once())
        try:
            self._loop.run_until_complete(task)
        finally:
            if not task.done():
                task.cancel()
                with suppress(asyncio.CancelledError):
                    self._loop.run_until_complete(task)
            self._loop.run_until_complete(self._loop.shutdown_asyncgens())
            self._loop.close()

    async def _run_once(self) -> None:
        timeout = aiohttp.ClientTimeout(total=60)
        async with aiohttp.ClientSession(timeout=timeout) as session:
            async with session.ws_connect(self.ws_url, heartbeat=20) as ws:
                self._force_logout_requested = False
                self._logged_out = False
                self.logger.info("mail websocket client connected: %s", self.ws_url)

                capabilities_msg = await ws.receive()
                if capabilities_msg.type != aiohttp.WSMsgType.TEXT:
                    raise RuntimeError("first message must be JSON-RPC capabilities")
                capabilities = json.loads(capabilities_msg.data)
                self.logger.info("capabilities=%s", capabilities)

                login_resp = await self._rpc_call(
                    ws,
                    rpc_id=1,
                    method="auth.login",
                    params={"account": self.account, "password": self.password},
                )
                cookie = str(login_resp.get("cookie", "")).strip()
                if not cookie:
                    raise RuntimeError("login success but cookie is empty")
                else:
                    self.logger.info(
                        "login success: account=%s, http://%s:%s/view/mail/folders?cookie=%s",
                        self.account,
                        self.server_host,
                        self.server_port,
                        cookie,
                    )


                confirm_resp = await self._rpc_call(
                    ws,
                    rpc_id=2,
                    method="auth.confirm",
                    params={"cookie": cookie},
                )
                self.logger.info("login confirmed, enabled_methods=%s", confirm_resp.get("enabled_methods"))

                token_resp = await self._rpc_call(
                    ws,
                    rpc_id=3,
                    method="outlook.token.acquire",
                    params={"cookie": cookie},
                )
                folders = token_resp.get("folders", [])
                self.logger.info(
                    "token_success=%s folder_count=%s",
                    token_resp.get("success"),
                    len(folders) if isinstance(folders, list) else 0,
                )
                self.logger.info("folders=%s", self._format_folders_for_log(folders))
                csv_path = self._save_folders_to_csv(folders)
                await self._sync_mode_num_counts(ws, cookie, csv_path)
                if await self._logout_if_requested(ws, cookie):
                    await ws.close()
                    self.logger.info("mail websocket client finished one-shot flow (forced logout)")
                    return

                await self._post_flow_pull_folders(ws, cookie)
                if await self._logout_if_requested(ws, cookie):
                    await ws.close()
                    self.logger.info("mail websocket client finished one-shot flow (forced logout)")
                    return

                await self._send_logout(ws, cookie, rpc_id=4)

                await ws.close()
                self.logger.info("mail websocket client finished one-shot flow")

    async def _post_flow_pull_folders(self, ws: aiohttp.ClientWebSocketResponse, cookie: str) -> None:
        for i in range(1, self.POST_FLOW_FOLDER_PULL_TIMES + 1):
            if self._stop_event.is_set():
                break
            await self._wait_with_server_push(ws, self.POST_FLOW_FOLDER_PULL_INTERVAL_SECONDS)
            if await self._logout_if_requested(ws, cookie):
                break
            try:
                token_resp = await self._rpc_call(
                    ws,
                    rpc_id=4000 + i,
                    method="outlook.token.acquire",
                    params={"cookie": cookie},
                )
                folders = token_resp.get("folders", [])
                self._save_folders_to_csv(folders)
                self.logger.info(
                    "post-flow folder pull %s/%s: success=%s folders=%s",
                    i,
                    self.POST_FLOW_FOLDER_PULL_TIMES,
                    token_resp.get("success"),
                    len(folders) if isinstance(folders, list) else 0,
                )
            except Exception:
                self.logger.exception(
                    "post-flow folder pull failed: round=%s/%s",
                    i,
                    self.POST_FLOW_FOLDER_PULL_TIMES,
                )
            if await self._logout_if_requested(ws, cookie):
                break

    async def _wait_with_server_push(self, ws: aiohttp.ClientWebSocketResponse, wait_seconds: int) -> None:
        deadline = time.time() + max(0, wait_seconds)
        while True:
            remaining = deadline - time.time()
            if remaining <= 0:
                return
            timeout = min(1.0, remaining)
            try:
                msg = await asyncio.wait_for(ws.receive(), timeout=timeout)
            except asyncio.TimeoutError:
                continue
            if msg.type != aiohttp.WSMsgType.TEXT:
                continue
            payload = json.loads(msg.data)
            if payload.get("jsonrpc") != "2.0":
                continue
            if "method" in payload:
                await self._handle_server_rpc_request(ws, payload)
                continue
            self.logger.warning("unexpected rpc payload while waiting: %s", payload)

    async def _logout_if_requested(self, ws: aiohttp.ClientWebSocketResponse, cookie: str) -> bool:
        if not self._force_logout_requested:
            return False
        await self._send_logout(ws, cookie, rpc_id=4999)
        return True

    async def _send_logout(self, ws: aiohttp.ClientWebSocketResponse, cookie: str, rpc_id: int) -> None:
        if self._logged_out:
            return
        logout_resp = await self._rpc_call(
            ws,
            rpc_id=rpc_id,
            method="auth.logout",
            params={"cookie": cookie},
        )
        self._logged_out = True
        self.logger.info("logout=%s", logout_resp)

    async def _rpc_call(
        self,
        ws: aiohttp.ClientWebSocketResponse,
        rpc_id: int,
        method: str,
        params: dict[str, Any],
    ) -> dict[str, Any]:
        request_unixtime_ms = int(time.time() * 1000)
        rpc_start = time.perf_counter()
        request = {
            "jsonrpc": "2.0",
            "id": rpc_id,
            "method": method,
            "params": params,
            "unixtime_ms": request_unixtime_ms,
        }
        await ws.send_json(request)
        while True:
            msg = await asyncio.wait_for(ws.receive(), timeout=self.RPC_TIMEOUT_SECONDS)
            if msg.type != aiohttp.WSMsgType.TEXT:
                raise RuntimeError(f"rpc response not text, method={method}")
            payload = json.loads(msg.data)
            if payload.get("jsonrpc") != "2.0":
                raise RuntimeError(f"invalid rpc response: {payload}")

            # 鏈嶅姟绔彲鑳藉湪浠绘剰鏃跺埢鍙嶅悜璇锋眰瀹㈡埛绔紝鍏堝鐞嗗悗缁х画绛夊緟鏈璋冪敤鍝嶅簲銆?            if "method" in payload:
                await self._handle_server_rpc_request(ws, payload)
                continue
            if payload.get("id") != rpc_id:
                self.logger.warning("skip unmatched rpc response: expect_id=%s payload=%s", rpc_id, payload)
                continue
            break

        echo_request_ms = payload.get("request_unixtime_ms")
        response_ms = payload.get("response_unixtime_ms")
        if not isinstance(response_ms, int):
            self.logger.warning("rpc response missing response_unixtime_ms, method=%s payload=%s", method, payload)
        if echo_request_ms != request_unixtime_ms:
            self.logger.warning(
                "rpc request_unixtime_ms mismatch, method=%s sent=%s recv=%s",
                method,
                request_unixtime_ms,
                echo_request_ms,
            )

        if "error" in payload:
            error = payload.get("error", {})
            raise RuntimeError(f"rpc error method={method}, code={error.get('code')}, message={error.get('message')}")

        result = payload.get("result", {})
        if not isinstance(result, dict):
            raise RuntimeError(f"rpc result must be object, method={method}")

        # 鎬ц兘瑙傚療锛氬鎴风鍙戝嚭 JSON-RPC 鍒版敹鍒版枃浠跺す鍒楄〃鐨勭鍒扮鑰楁椂锛堥噸鐐硅娴?outlook.token.acquire锛夈€?        if method == "outlook.token.acquire":
            elapsed_ms = int((time.perf_counter() - rpc_start) * 1000)
            folders = result.get("folders", [])
            folder_count = len(folders) if isinstance(folders, list) else 0
            self.logger.info(
                "perf rpc acquire: method=%s elapsed_ms=%s folder_count=%s request_unixtime_ms=%s response_unixtime_ms=%s",
                method,
                elapsed_ms,
                folder_count,
                request_unixtime_ms,
                response_ms,
            )
        return result

    async def _handle_server_rpc_request(
        self,
        ws: aiohttp.ClientWebSocketResponse,
        payload: dict[str, Any],
    ) -> None:
        method = str(payload.get("method", "")).strip()
        rpc_id = payload.get("id")
        params = payload.get("params", {})
        params = params if isinstance(params, dict) else {}
        if not isinstance(rpc_id, int):
            return

        if method == self.RPC_METHOD_CLIENT_FORCE_LOGOUT:
            cookie = str(params.get("cookie", "")).strip()
            self._force_logout_requested = True
            self.logger.info("force logout requested by server: account=%s cookie=%s", self.account, cookie)
            await ws.send_json(
                {
                    "jsonrpc": "2.0",
                    "id": rpc_id,
                    "result": {
                        "success": True,
                        "message": "force logout accepted",
                        "account": self.account,
                    },
                    "request_unixtime_ms": payload.get("unixtime_ms"),
                    "response_unixtime_ms": int(time.time() * 1000),
                }
            )
            return

        if method == self.RPC_METHOD_LOCAL_FOLDER_LIST:
            try:
                rows, csv_path = self._load_local_folder_rows()
                await ws.send_json(
                    {
                        "jsonrpc": "2.0",
                        "id": rpc_id,
                        "result": {
                            "success": True,
                            "account": self.account,
                            "csv_path": str(csv_path),
                            "folders": rows,
                        },
                        "request_unixtime_ms": payload.get("unixtime_ms"),
                        "response_unixtime_ms": int(time.time() * 1000),
                    }
                )
            except Exception as exc:
                self.logger.exception("handle server rpc request failed: method=%s", method)
                await ws.send_json(
                    {
                        "jsonrpc": "2.0",
                        "id": rpc_id,
                        "error": {"code": -32090, "message": str(exc)},
                        "request_unixtime_ms": payload.get("unixtime_ms"),
                        "response_unixtime_ms": int(time.time() * 1000),
                    }
                )
            return

        if method == self.RPC_METHOD_LOCAL_TITLE_LIST:
            folder_name = str(params.get("folder_name", "")).strip()
            try:
                rows, csv_path = self._load_local_title_rows(folder_name)
                await ws.send_json(
                    {
                        "jsonrpc": "2.0",
                        "id": rpc_id,
                        "result": {
                            "success": True,
                            "account": self.account,
                            "folder_name": folder_name,
                            "csv_path": str(csv_path),
                            "titles": rows,
                        },
                        "request_unixtime_ms": payload.get("unixtime_ms"),
                        "response_unixtime_ms": int(time.time() * 1000),
                    }
                )
            except Exception as exc:
                self.logger.exception("handle server rpc request failed: method=%s", method)
                await ws.send_json(
                    {
                        "jsonrpc": "2.0",
                        "id": rpc_id,
                        "error": {"code": -32090, "message": str(exc)},
                        "request_unixtime_ms": payload.get("unixtime_ms"),
                        "response_unixtime_ms": int(time.time() * 1000),
                    }
                )
            return

        await ws.send_json(
            {
                "jsonrpc": "2.0",
                "id": rpc_id,
                "error": {"code": -32601, "message": f"Method not found: {method}"},
                "request_unixtime_ms": payload.get("unixtime_ms"),
                "response_unixtime_ms": int(time.time() * 1000),
            }
        )

    @staticmethod
    def _format_folders_for_log(folders: Any) -> list[dict[str, Any]]:
        if not isinstance(folders, list):
            return []
        normalized: list[dict[str, Any]] = []
        for item in folders:
            if not isinstance(item, dict):
                continue
            normalized.append(
                {
                    "name": str(item.get("name", "")),
                    "flags": [str(x) for x in item.get("flags", [])] if isinstance(item.get("flags", []), list) else [],
                }
            )
        return normalized

    def _save_folders_to_csv(self, folders: Any) -> Path:
        rows = sorted(self._format_folders_for_log(folders), key=lambda x: x.get("name", "").lower())
        output = self._folders_csv_path()
        output.parent.mkdir(parents=True, exist_ok=True)
        now_ms = int(time.time() * 1000)

        existing_by_name = self._load_existing_rows_by_name(output)
        self._set_file_writable(output)
        try:
            with output.open("w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=self.FOLDER_CSV_FIELDS)
                writer.writeheader()
                for item in rows:
                    existed = existing_by_name.get(item["name"], {})
                    writer.writerow(
                        {
                            "unixtime_ms": now_ms,
                            "name": item["name"],
                            "flags": json.dumps(item["flags"], ensure_ascii=False),
                            "mode": existed.get("mode", ""),
                            "current_count": existed.get("current_count", ""),
                            "online_count": existed.get("online_count", ""),
                            "current_unixtime_ms": existed.get("current_unixtime_ms", ""),
                            "update_unixtime_ms": existed.get("update_unixtime_ms", ""),
                        }
                    )
        finally:
            self._set_file_readonly(output)
        self.logger.info("folders csv saved: %s", output.resolve())
        return output.resolve()

    def _folders_csv_path(self) -> Path:
        safe_prefix = self._account_file_prefix()
        return Path("result") / f"{safe_prefix}_folders.csv"

    def _account_file_prefix(self) -> str:
        account = (self.account or "").strip()
        return re.sub(r"[\\/:*?\"<>|]", "_", account) or "mail"

    def _title_csv_path(self, folder_name: str) -> Path:
        safe_prefix = self._account_file_prefix()
        safe_folder = re.sub(r"[\\/:*?\"<>|]", "_", folder_name.strip()) or "folder"
        return Path("result") / f"{safe_prefix}_{safe_folder}.csv"

    def _load_local_folder_rows(self) -> tuple[list[dict[str, str]], Path]:
        csv_path = self._folders_csv_path().resolve()
        if not csv_path.exists():
            return [], csv_path
        with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
            rows = list(csv.DictReader(f))
        normalized: list[dict[str, str]] = []
        for row in rows:
            normalized.append(
                {
                    "name": str(row.get("name", "")).strip(),
                    "flags": str(row.get("flags", "")).strip(),
                    "mode": str(row.get("mode", "")).strip(),
                    "current_count": str(row.get("current_count", "")).strip(),
                    "online_count": str(row.get("online_count", "")).strip(),
                }
            )
        return normalized, csv_path

    def _load_local_title_rows(self, folder_name: str) -> tuple[list[dict[str, str]], Path]:
        csv_path = self._title_csv_path(folder_name).resolve()
        if not csv_path.exists():
            return [], csv_path
        with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
            rows = list(csv.DictReader(f))
        normalized: list[dict[str, str]] = []
        for row in rows:
            normalized.append(
                {
                    "mail_id": str(row.get("mail_id", "")).strip(),
                    "uid": str(row.get("uid", "")).strip(),
                    "message_id": str(row.get("message_id", "")).strip(),
                    "sender": str(row.get("sender", "")).strip(),
                    "title": str(row.get("title", "")).strip(),
                    "received_at": str(row.get("received_at", "")).strip(),
                    "received_unixtime_ms": str(row.get("received_unixtime_ms", "")).strip(),
                }
            )
        return normalized, csv_path

    @staticmethod
    def _load_existing_rows_by_name(path: Path) -> dict[str, dict[str, Any]]:
        if not path.exists():
            return {}
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            rows = list(reader)
        result: dict[str, dict[str, Any]] = {}
        for idx, row in enumerate(rows):
            name = str(row.get("name", "")).strip()
            if name:
                result[name] = row
        return result

    async def _sync_mode_num_counts(self, ws: aiohttp.ClientWebSocketResponse, cookie: str, csv_path: Path) -> None:
        pending_notifications: list[dict[str, Any]] = []
        self._set_file_writable(csv_path)
        try:
            with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
                rows = list(csv.DictReader(f))

            changed = False
            for idx, row in enumerate(rows):
                mode = str(row.get("mode", "")).strip().lower()
                folder_name = str(row.get("name", "")).strip()
                if not folder_name:
                    continue

                if mode not in {"num", "title", "base64a"}:
                    continue

                current_raw = str(row.get("current_count", "")).strip()
                try:
                    request_count = int(current_raw) if current_raw else 0
                except ValueError:
                    request_count = 0

                rpc_result = await self._rpc_call(
                    ws,
                    rpc_id=10 + idx,
                    method="mail.folder.count",
                    params={
                        "cookie": cookie,
                        "folder_name": folder_name,
                        "current_count": request_count,
                    },
                )

                if not rpc_result.get("success"):
                    continue

                online_count = int(rpc_result.get("folder_count", 0))
                update_ms = int(rpc_result.get("update_unixtime_ms", int(time.time() * 1000)))
                row["online_count"] = str(online_count)
                row["current_count"] = str(online_count)
                row["current_unixtime_ms"] = str(update_ms)
                row["update_unixtime_ms"] = str(update_ms)
                changed = True
                self.logger.info(
                    "folder count synced: folder=%s request_count=%s online_count=%s",
                    folder_name,
                    request_count,
                    online_count,
                )

                if mode in {"title", "base64a"}:
                    existing_title_rows = self._load_existing_title_rows_for_merge(folder_name)
                    local_count = len(existing_title_rows)
                    known_max_uid = self._extract_known_max_uid(existing_title_rows)
                    incremental_count = max(0, online_count - local_count)
                    self.logger.info(
                        "title incremental plan: folder=%s mode=%s online_count=%s local_count=%s known_max_uid=%s incremental_count=%s",
                        folder_name,
                        mode,
                        online_count,
                        local_count,
                        known_max_uid,
                        incremental_count,
                    )
                    if incremental_count <= 0:
                        if local_count > online_count:
                            # Online count reduced (e.g. deleted/moved mails), trim local CSV to stay consistent.
                            trimmed_count = self._save_titles_to_csv(
                                folder_name=folder_name,
                                titles=[],
                                expected_count=online_count,
                                merge_existing=True,
                                existing_rows=existing_title_rows,
                            )
                            self.logger.info(
                                "title incremental trim: folder=%s mode=%s local_count=%s online_count=%s trimmed_count=%s",
                                folder_name,
                                mode,
                                local_count,
                                online_count,
                                trimmed_count,
                            )
                        self.logger.info(
                            "title incremental skip: folder=%s mode=%s reason=no_new_mail",
                            folder_name,
                            mode,
                        )
                        continue
                    title_method = "title.base64a" if mode == "base64a" else "title"
                    title_params: dict[str, Any] = {
                        "cookie": cookie,
                        "folder_name": folder_name,
                        "incremental_count": incremental_count,
                    }
                    if known_max_uid is not None:
                        title_params["known_max_uid"] = known_max_uid
                    t_title_rpc_start = time.perf_counter()
                    title_result = await self._rpc_call(
                        ws,
                        rpc_id=1000 + idx,
                        method=title_method,
                        params=title_params,
                    )
                    title_rpc_ms = int((time.perf_counter() - t_title_rpc_start) * 1000)
                    if title_result.get("success"):
                        titles = title_result.get("titles", [])
                        t_title_save_start = time.perf_counter()
                        merged_count = self._save_titles_to_csv(
                            folder_name=folder_name,
                            titles=titles,
                            expected_count=online_count,
                            merge_existing=True,
                            existing_rows=existing_title_rows,
                        )
                        title_save_ms = int((time.perf_counter() - t_title_save_start) * 1000)
                        increment_fetched = len(titles) if isinstance(titles, list) else 0
                        pending_notifications.append(
                            self._build_title_notification_item(
                                folder_name=folder_name,
                                titles=titles,
                                expected_count=online_count,
                            )
                        )
                        self.logger.info(
                            "folder titles synced: folder=%s mode=%s increment_fetched=%s merged_count=%s",
                            folder_name,
                            mode,
                            increment_fetched,
                            merged_count,
                        )
                        self.logger.info(
                            "perf title incremental: folder=%s mode=%s online_count=%s local_count=%s incremental_count=%s fetched=%s merged_count=%s rpc_ms=%s save_ms=%s total_ms=%s",
                            folder_name,
                            mode,
                            online_count,
                            local_count,
                            incremental_count,
                            increment_fetched,
                            merged_count,
                            title_rpc_ms,
                            title_save_ms,
                            title_rpc_ms + title_save_ms,
                        )

            if not changed:
                return

            with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=self.FOLDER_CSV_FIELDS)
                writer.writeheader()
                for row in rows:
                    writer.writerow({key: row.get(key, "") for key in self.FOLDER_CSV_FIELDS})
        finally:
            self._set_file_readonly(csv_path)
        self.logger.info("folders csv synced by mode=num/title/base64a: %s", csv_path)
        if pending_notifications:
            await self._send_title_notifications_batch(ws, cookie, pending_notifications)

    def _save_titles_to_csv(
        self,
        folder_name: str,
        titles: Any,
        expected_count: int,
        merge_existing: bool = False,
        existing_rows: list[dict[str, Any]] | None = None,
    ) -> int:
        output = self._title_csv_path(folder_name)
        output.parent.mkdir(parents=True, exist_ok=True)

        rows: list[dict[str, Any]] = []
        if merge_existing:
            rows.extend(existing_rows if existing_rows is not None else self._load_existing_title_rows_for_merge(folder_name))

        now_ms = int(time.time() * 1000)
        if isinstance(titles, list):
            for item in titles:
                if not isinstance(item, dict):
                    continue
                rows.append(
                    {
                        "unixtime_ms": now_ms,
                        "mail_id": str(item.get("mail_id", "")),
                        "uid": str(item.get("uid", "")),
                        "message_id": str(item.get("message_id", "")),
                        "title": str(item.get("title", "")),
                        "sender": str(item.get("sender", "")),
                        "received_at": str(item.get("received_at", "")),
                        "received_unixtime_ms": int(item.get("received_unixtime_ms", 0) or 0),
                        "Base64A": str(item.get("Base64A", "") or ""),
                    }
                )

        merged_by_key: dict[str, dict[str, Any]] = {}
        for row in rows:
            merged_by_key[self._title_row_key(row)] = row
        rows = list(merged_by_key.values())
        rows.sort(
            key=lambda x: (
                self._safe_int(x.get("received_unixtime_ms", 0)),
                self._safe_int(x.get("uid", 0)),
            )
        )
        if expected_count >= 0 and len(rows) > expected_count:
            rows = rows[-expected_count:] if expected_count > 0 else []

        self._set_file_writable(output)
        try:
            with output.open("w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=self.TITLE_CSV_FIELDS, quoting=csv.QUOTE_ALL)
                writer.writeheader()
                for row in rows:
                    writer.writerow({key: row.get(key, "") for key in self.TITLE_CSV_FIELDS})
        finally:
            self._set_file_readonly(output)

        sender_missing = sum(1 for row in rows if not str(row.get("sender", "")).strip())
        self.logger.info(
            "title csv stats: folder=%s rows=%s sender_missing=%s",
            folder_name,
            len(rows),
            sender_missing,
        )
        self.logger.info("title csv saved: %s", output.resolve())
        return len(rows)

    def _load_existing_title_rows_for_merge(self, folder_name: str) -> list[dict[str, Any]]:
        output = self._title_csv_path(folder_name)
        if not output.exists():
            return []
        with output.open("r", encoding="utf-8-sig", newline="") as f:
            rows = list(csv.DictReader(f))

        normalized: list[dict[str, Any]] = []
        for row in rows:
            normalized.append(
                {
                    "mail_id": str(row.get("mail_id", "")),
                    "uid": str(row.get("uid", "")),
                    "message_id": str(row.get("message_id", "")),
                    "sender": str(row.get("sender", "")),
                    "title": str(row.get("title", "")),
                    "received_at": str(row.get("received_at", "")),
                    "received_unixtime_ms": self._safe_int(row.get("received_unixtime_ms", 0)),
                    "unixtime_ms": self._safe_int(row.get("unixtime_ms", 0)),
                    "Base64A": str(row.get("Base64A", "") or ""),
                }
            )
        return normalized

    @staticmethod
    def _extract_known_max_uid(rows: list[dict[str, Any]]) -> int | None:
        max_uid: int | None = None
        for row in rows:
            uid = InternalWSClient._safe_int(row.get("uid", 0))
            if uid <= 0:
                continue
            if max_uid is None or uid > max_uid:
                max_uid = uid
        return max_uid

    @staticmethod
    def _title_row_key(row: dict[str, Any]) -> str:
        uid = str(row.get("uid", "")).strip()
        if uid:
            return f"uid:{uid}"
        message_id = str(row.get("message_id", "")).strip()
        if message_id:
            return f"message_id:{message_id}"
        mail_id = str(row.get("mail_id", "")).strip()
        if mail_id:
            return f"mail_id:{mail_id}"
        return f"fallback:{str(row.get('received_at', '')).strip()}|{str(row.get('title', '')).strip()}"

    @staticmethod
    def _safe_int(value: Any, default: int = 0) -> int:
        try:
            return int(value)
        except (TypeError, ValueError):
            return default


    def _build_title_notification_item(self, folder_name: str, titles: Any, expected_count: int) -> dict[str, Any]:
        rows: list[dict[str, Any]] = []
        if isinstance(titles, list):
            for item in titles:
                if not isinstance(item, dict):
                    continue
                received_ms = int(item.get("received_unixtime_ms", 0) or 0)
                rows.append(
                    {
                        "title": str(item.get("title", "")).strip() or "(No Subject)",
                        "sender": str(item.get("sender", "")).strip() or "(Unknown Sender)",
                        "received_at": str(item.get("received_at", "")).strip() or "(Unknown Time)",
                        "received_unixtime_ms": received_ms,
                    }
                )

        rows.sort(key=lambda x: x["received_unixtime_ms"], reverse=True)
        return {
            "folder_name": folder_name,
            "expected_count": expected_count,
            "total_count": len(rows),
            "rows": rows[: self.FEISHU_TITLE_LIMIT],
        }

    @staticmethod
    def _truncate_text(text: str, limit: int) -> str:
        clean = " ".join((text or "").split())
        if len(clean) <= limit:
            return clean
        return f"{clean[:limit]}..."

    async def _send_title_notifications_batch(
        self,
        ws: aiohttp.ClientWebSocketResponse,
        cookie: str,
        notifications: list[dict[str, Any]],
    ) -> None:
        if not notifications:
            return

        lines: list[str] = [
            f"Account: {self.account}",
            f"Folders synced this round: {len(notifications)}",
            "",
        ]

        for item in notifications:
            folder_name = str(item.get("folder_name", ""))
            expected_count = int(item.get("expected_count", 0))
            total_count = int(item.get("total_count", 0))
            rows = item.get("rows", [])
            rows = rows if isinstance(rows, list) else []
            lines.append(f"[{folder_name}] online={expected_count} fetched={total_count}")
            show_rows = rows[: self.FEISHU_PER_FOLDER_LIMIT]
            for idx, row in enumerate(show_rows, start=1):
                sender = self._truncate_text(str(row.get("sender", "(Unknown Sender)")), 30)
                subject = self._truncate_text(str(row.get("title", "(No Subject)")), 80)
                received_at = str(row.get("received_at", "(Unknown Time)"))
                lines.append(f"{idx}. [{received_at}] {sender} | {subject}")
            if total_count > self.FEISHU_PER_FOLDER_LIMIT:
                lines.append(f"... and {total_count - self.FEISHU_PER_FOLDER_LIMIT} more")
            lines.append("")

        body = "\n".join(lines).strip()
        if len(body) > self.FEISHU_BODY_MAX_CHARS:
            body = f"{body[: self.FEISHU_BODY_MAX_CHARS - 16]}\n...truncated"

        title = f"Outlook Title Sync/{self.account}"
        try:
            rpc_result = await self._rpc_call(
                ws,
                rpc_id=900000 + len(notifications),
                method="feishu.notify",
                params={
                    "cookie": cookie,
                    "body": body,
                    "title": title,
                },
            )
            results = rpc_result.get("results", {})
            results = results if isinstance(results, dict) else {}
            success_count = int(rpc_result.get("success_count", 0))
            self.logger.info(
                "feishu title notify sent: folders=%s targets=%s success=%s",
                len(notifications),
                len(results),
                success_count,
            )
        except Exception:
            self.logger.exception("send feishu title notification failed")

    def _set_file_writable(self, path: Path) -> None:
        if not path.exists():
            return
        try:
            path.chmod(stat.S_IWRITE | stat.S_IREAD)
            self.logger.debug("set writable: %s", path)
        except OSError:
            self.logger.warning("failed to clear readonly attribute: %s", path, exc_info=True)

    def _set_file_readonly(self, path: Path) -> None:
        if not path.exists():
            return
        try:
            path.chmod(stat.S_IREAD)
            self.logger.debug("set readonly: %s", path)
        except OSError:
            self.logger.warning("failed to set readonly attribute: %s", path, exc_info=True)

    @staticmethod
    def _ensure_csv_field_limit() -> None:
        current = csv.field_size_limit()
        target = 16 * 1024 * 1024
        if current >= target:
            return
        try:
            csv.field_size_limit(target)
        except OverflowError:
            csv.field_size_limit(current)
