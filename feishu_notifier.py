from pathlib import Path
from typing import Dict, List, Optional

import aiohttp
from public.csv_ops import CsvOps
from public.feishu_ops import FeishuOps


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
            rows = CsvOps.read_rows(self.config_file, encoding="utf-8-sig")
            for row in rows:
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
        return FeishuOps.build_message(body=body, title=title, mode=mode)

    async def _send_to_webhook(self, session: aiohttp.ClientSession, url: str, message: dict) -> bool:
        return await FeishuOps.send_to_webhook(session=session, url=url, message=message, logger=self.logger)

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
