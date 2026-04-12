import asyncio
from contextlib import suppress
import json
import threading
from typing import Any

import aiohttp


class InternalWSClient:
    """邮件模块 WebSocket 客户端，按 JSON-RPC 2.0 流程完成一次登录与查询。"""

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
                self.logger.info(
                    "token_success=%s mailbox=%s inbox_total=%s",
                    token_resp.get("success"),
                    token_resp.get("mailbox"),
                    token_resp.get("inbox_total"),
                )

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
        request = {
            "jsonrpc": "2.0",
            "id": rpc_id,
            "method": method,
            "params": params,
        }
        await ws.send_json(request)
        msg = await ws.receive()
        if msg.type != aiohttp.WSMsgType.TEXT:
            raise RuntimeError(f"rpc response not text, method={method}")

        payload = json.loads(msg.data)
        if payload.get("jsonrpc") != "2.0" or payload.get("id") != rpc_id:
            raise RuntimeError(f"invalid rpc response: {payload}")

        if "error" in payload:
            error = payload.get("error", {})
            raise RuntimeError(f"rpc error method={method}, code={error.get('code')}, message={error.get('message')}")

        result = payload.get("result", {})
        if not isinstance(result, dict):
            raise RuntimeError(f"rpc result must be object, method={method}")
        return result
