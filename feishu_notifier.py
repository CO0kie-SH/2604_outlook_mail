import csv
from pathlib import Path
from typing import Dict, List, Optional

import aiohttp


class FeishuNotifier:
    def __init__(self, config_dir: str = "config", logger=None):
        self.config_dir = Path(config_dir)
        self.config_file = self.config_dir / "FeiShu.csv"
        self.logger = logger
        self.configs = self._load_configs()

    def _load_configs(self) -> List[Dict[str, str]]:
        if not self.config_file.exists():
            if self.logger:
                self.logger.warning("飞书配置文件不存在: %s", self.config_file)
            return []

        configs: List[Dict[str, str]] = []
        try:
            with self.config_file.open("r", encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    tag = str(row.get("tag", "")).strip()
                    url = str(row.get("url", "")).strip()
                    mode = str(row.get("mode", "")).strip().lower()
                    if not tag or not url:
                        continue
                    configs.append({"tag": tag, "url": url, "mode": mode or "text"})
            if self.logger:
                self.logger.info("加载飞书配置 %s 条", len(configs))
        except Exception as exc:
            if self.logger:
                self.logger.error("加载飞书配置失败: %s", exc, exc_info=True)
        return configs

    @staticmethod
    def _build_message(body: str, title: Optional[str] = None, mode: str = "text") -> dict:
        if mode == "text" or title is None:
            return {
                "msg_type": "text",
                "content": {"text": body},
            }
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

    async def _send_to_webhook(self, session: aiohttp.ClientSession, url: str, message: dict) -> bool:
        try:
            async with session.post(url, json=message) as response:
                text = await response.text()
                if response.status != 200:
                    if self.logger:
                        self.logger.error("飞书消息发送失败，状态码=%s body=%s", response.status, text)
                    return False
                try:
                    result = await response.json()
                except Exception:
                    result = {}
                if result.get("StatusCode") == 0:
                    if self.logger:
                        self.logger.info("飞书消息发送成功")
                    return True
                if self.logger:
                    self.logger.error("飞书消息发送失败: %s", result or text)
                return False
        except Exception as exc:
            if self.logger:
                self.logger.error("飞书消息发送异常: %s", exc, exc_info=True)
            return False

    async def send_message(
        self,
        body: str,
        title: Optional[str] = None,
        tag: Optional[str] = None,
    ) -> Dict[str, bool]:
        results: Dict[str, bool] = {}

        if not self.configs:
            if self.logger:
                self.logger.warning("无飞书配置，跳过发送")
            return results

        timeout = aiohttp.ClientTimeout(total=30)
        async with aiohttp.ClientSession(timeout=timeout) as session:
            for config in self.configs:
                config_tag = config["tag"]
                config_url = config["url"]
                config_mode = config["mode"]

                if tag and config_tag != tag:
                    continue

                if config_mode == "none":
                    results[config_tag] = True
                    if self.logger:
                        self.logger.info("飞书机器人 [%s] 模式为 none，跳过发送", config_tag)
                    continue

                mode = "text" if config_mode == "text" else "post"
                message = self._build_message(body=body, title=title, mode=mode)
                results[config_tag] = await self._send_to_webhook(session, config_url, message)

        return results

    async def send_to_all(self, body: str, title: Optional[str] = None) -> Dict[str, bool]:
        return await self.send_message(body=body, title=title)

    async def send_to_tag(self, tag: str, body: str, title: Optional[str] = None) -> bool:
        results = await self.send_message(body=body, title=title, tag=tag)
        return results.get(tag, False)


async def send_feishu_message(
    logger,
    v_body: str,
    v_title: Optional[str] = None,
    tag: Optional[str] = None,
) -> Dict[str, bool]:
    notifier = FeishuNotifier(logger=logger)
    return await notifier.send_message(body=v_body, title=v_title, tag=tag)

