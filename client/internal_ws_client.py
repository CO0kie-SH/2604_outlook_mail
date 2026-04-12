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

from feishu_notifier import send_feishu_message


class InternalWSClient:
    """邮件模块 WebSocket 客户端，按 JSON-RPC 2.0 流程完成一次登录与查询。"""
    RPC_TIMEOUT_SECONDS = 20
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
        "sender",
        "title",
        "received_at",
        "received_unixtime_ms",
        "unixtime_ms",
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
                self.logger.info("mail websocket client connected: %s", self.ws_url)

                # 1) 接收服务端首次下发的接口列表。
                capabilities_msg = await ws.receive()
                if capabilities_msg.type != aiohttp.WSMsgType.TEXT:
                    raise RuntimeError("first message must be JSON-RPC capabilities")
                capabilities = json.loads(capabilities_msg.data)
                self.logger.info("capabilities=%s", capabilities)

                # 2) 请求登录并拿到 cookie。
                login_resp = await self._rpc_call(
                    ws,
                    rpc_id=1,
                    method="auth.login",
                    params={"account": self.account, "password": self.password},
                )
                cookie = str(login_resp.get("cookie", "")).strip()
                if not cookie:
                    raise RuntimeError("login success but cookie is empty")

                # 3) 发送 cookie 进行登录确认，获取已登录可用接口。
                confirm_resp = await self._rpc_call(
                    ws,
                    rpc_id=2,
                    method="auth.confirm",
                    params={"cookie": cookie},
                )
                self.logger.info("login confirmed, enabled_methods=%s", confirm_resp.get("enabled_methods"))

                # 4) 请求 outlook token 信息，并获取一次 INBOX 邮件总数。
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

                # 5) 首次拿到邮件数后退出登录，不再重复登录。
                logout_resp = await self._rpc_call(
                    ws,
                    rpc_id=4,
                    method="auth.logout",
                    params={"cookie": cookie},
                )
                self.logger.info("logout=%s", logout_resp)

                await ws.close()
                self.logger.info("mail websocket client finished one-shot flow")

    async def _rpc_call(
        self,
        ws: aiohttp.ClientWebSocketResponse,
        rpc_id: int,
        method: str,
        params: dict[str, Any],
    ) -> dict[str, Any]:
        request_unixtime_ms = int(time.time() * 1000)
        request = {
            "jsonrpc": "2.0",
            "id": rpc_id,
            "method": method,
            "params": params,
            "unixtime_ms": request_unixtime_ms,
        }
        await ws.send_json(request)
        msg = await asyncio.wait_for(ws.receive(), timeout=self.RPC_TIMEOUT_SECONDS)
        if msg.type != aiohttp.WSMsgType.TEXT:
            raise RuntimeError(f"rpc response not text, method={method}")

        payload = json.loads(msg.data)
        if payload.get("jsonrpc") != "2.0" or payload.get("id") != rpc_id:
            raise RuntimeError(f"invalid rpc response: {payload}")

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
        return result

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
        account = (self.account or "").strip()
        local_part = account.split("@", 1)[0].strip() if account else ""
        safe_prefix = re.sub(r"[\\/:*?\"<>|]", "_", local_part) or "mail"
        output = Path("config") / f"{safe_prefix}_folders.csv"
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

                if mode not in {"num", "title"}:
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

                if mode == "title":
                    title_result = await self._rpc_call(
                        ws,
                        rpc_id=1000 + idx,
                        method="title",
                        params={
                            "cookie": cookie,
                            "folder_name": folder_name,
                        },
                    )
                    if title_result.get("success"):
                        titles = title_result.get("titles", [])
                        self._save_titles_to_csv(folder_name, titles, online_count)
                        pending_notifications.append(
                            self._build_title_notification_item(
                                folder_name=folder_name,
                                titles=titles,
                                expected_count=online_count,
                            )
                        )
                        self.logger.info(
                            "folder titles synced: folder=%s title_count=%s expected_count=%s",
                            folder_name,
                            len(titles) if isinstance(titles, list) else 0,
                            online_count,
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
        self.logger.info("folders csv synced by mode=num/title: %s", csv_path)
        if pending_notifications:
            await self._send_title_notifications_batch(pending_notifications)

    def _save_titles_to_csv(self, folder_name: str, titles: Any, expected_count: int) -> None:
        account = (self.account or "").strip()
        local_part = account.split("@", 1)[0].strip() if account else ""
        safe_prefix = re.sub(r"[\\/:*?\"<>|]", "_", local_part) or "mail"
        safe_folder = re.sub(r"[\\/:*?\"<>|]", "_", folder_name.strip()) or "folder"
        output = Path("config") / f"{safe_prefix}_{safe_folder}.csv"
        output.parent.mkdir(parents=True, exist_ok=True)

        rows: list[dict[str, Any]] = []
        if isinstance(titles, list):
            for item in titles:
                if not isinstance(item, dict):
                    continue
                rows.append(
                    {
                        "unixtime_ms": int(time.time() * 1000),
                        "mail_id": str(item.get("mail_id", "")),
                        "title": str(item.get("title", "")),
                        "sender": str(item.get("sender", "")),
                        "received_at": str(item.get("received_at", "")),
                        "received_unixtime_ms": int(item.get("received_unixtime_ms", 0) or 0),
                    }
                )
        # 按送达时间升序（旧 -> 新）输出。
        rows.sort(key=lambda x: x["received_unixtime_ms"])
        if expected_count >= 0:
            original_count = len(rows)
            rows = rows[:expected_count]
            if original_count != expected_count:
                self.logger.warning(
                    "title count mismatch before trim: folder=%s expected=%s actual=%s",
                    folder_name,
                    expected_count,
                    original_count,
                )

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

    def _build_title_notification_item(self, folder_name: str, titles: Any, expected_count: int) -> dict[str, Any]:
        rows: list[dict[str, Any]] = []
        if isinstance(titles, list):
            for item in titles:
                if not isinstance(item, dict):
                    continue
                received_ms = int(item.get("received_unixtime_ms", 0) or 0)
                rows.append(
                    {
                        "title": str(item.get("title", "")).strip() or "(无主题)",
                        "sender": str(item.get("sender", "")).strip() or "(未知发件人)",
                        "received_at": str(item.get("received_at", "")).strip() or "(未知时间)",
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

    async def _send_title_notifications_batch(self, notifications: list[dict[str, Any]]) -> None:
        if not notifications:
            return

        lines: list[str] = [
            f"账号: {self.account}",
            f"本轮同步文件夹: {len(notifications)}",
            "",
        ]

        for item in notifications:
            folder_name = str(item.get("folder_name", ""))
            expected_count = int(item.get("expected_count", 0))
            total_count = int(item.get("total_count", 0))
            rows = item.get("rows", [])
            rows = rows if isinstance(rows, list) else []
            lines.append(f"【{folder_name}】在线数量: {expected_count} | 抓取数量: {total_count}")
            show_rows = rows[: self.FEISHU_PER_FOLDER_LIMIT]
            for idx, row in enumerate(show_rows, start=1):
                sender = self._truncate_text(str(row.get("sender", "(未知发件人)")), 30)
                subject = self._truncate_text(str(row.get("title", "(无主题)")), 80)
                received_at = str(row.get("received_at", "(未知时间)"))
                lines.append(f"{idx}. [{received_at}] {sender} | {subject}")
            if total_count > self.FEISHU_PER_FOLDER_LIMIT:
                lines.append(f"... 其余 {total_count - self.FEISHU_PER_FOLDER_LIMIT} 封省略")
            lines.append("")

        body = "\n".join(lines).strip()
        if len(body) > self.FEISHU_BODY_MAX_CHARS:
            body = f"{body[: self.FEISHU_BODY_MAX_CHARS - 16]}\n...内容已截断"

        title = f"Outlook标题抓取更新/{self.account}"
        try:
            results = await send_feishu_message(self.logger, body, v_title=title)
            success_count = sum(1 for ok in results.values() if ok)
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
