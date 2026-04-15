import logging
from typing import Optional

import aiohttp


class FeishuOps:
    logger = logging.getLogger(__name__)

    @staticmethod
    def build_message(body: str, title: Optional[str] = None, mode: str = "text") -> dict:
        FeishuOps.logger.info("build feishu message: mode=%s title=%s body_chars=%s", mode, bool(title), len(body))
        if mode == "text" or title is None:
            return {"msg_type": "text", "content": {"text": body}}
        return {
            "msg_type": "post",
            "content": {
                "post": {
                    "zh-CN": {
                        "title": title,
                        "content": [[{"tag": "text", "text": body}]],
                    }
                }
            },
        }

    @staticmethod
    async def send_to_webhook(
        session: aiohttp.ClientSession,
        url: str,
        message: dict,
        logger=None,
    ) -> bool:
        active_logger = logger or FeishuOps.logger
        try:
            active_logger.info("send feishu webhook: url=%s msg_type=%s", url, message.get("msg_type"))
            async with session.post(url, json=message) as response:
                text = await response.text()
                if response.status != 200:
                    active_logger.error("飞书消息发送失败，状态码=%s body=%s", response.status, text)
                    return False
                try:
                    result = await response.json()
                except Exception:
                    result = {}
                if result.get("StatusCode") == 0:
                    active_logger.info("飞书消息发送成功")
                    return True
                active_logger.error("飞书消息发送失败: %s", result or text)
                return False
        except Exception as exc:
            active_logger.error("飞书消息发送异常: %s", exc, exc_info=True)
            return False
