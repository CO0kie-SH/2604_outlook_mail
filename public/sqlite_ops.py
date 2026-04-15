import hashlib
import logging
import re
import time
from pathlib import Path
from typing import Any

import aiosqlite
import sqlite3


class Base64ASqliteStore:
    def __init__(self, db_path: Path, table_name: str = "mail_base64a"):
        self.db_path = db_path.resolve()
        self.table_name = self.normalize_identifier(table_name, fallback="mail_base64a")
        self._table_sql = f'"{self.table_name}"'
        self._index_sql = self.normalize_identifier(f"idx_{self.table_name}_folder_uid", fallback="idx_mail_base64a_folder_uid")
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def normalize_identifier(value: str, fallback: str = "table") -> str:
        text = (value or "").strip()
        text = re.sub(r"[^0-9A-Za-z_]", "_", text)
        text = re.sub(r"_+", "_", text).strip("_")
        if not text:
            text = fallback
        if text[0].isdigit():
            text = f"t_{text}"
        return text

    @staticmethod
    def build_account_table_name(account: str) -> str:
        return f"base64A_{Base64ASqliteStore.normalize_identifier(account, fallback='mail')}"

    @staticmethod
    def message_id_md5(message_id: str) -> str:
        value = (message_id or "").strip()
        if not value:
            return ""
        return hashlib.md5(value.encode("utf-8")).hexdigest()

    async def upsert_records(
        self,
        account: str,
        folder_name: str,
        records: list[dict[str, Any]],
    ) -> dict[str, str]:
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        now_ms = int(time.time() * 1000)
        md5_by_message_id: dict[str, str] = {}
        self.logger.info(
            "sqlite upsert start: db=%s table=%s account=%s folder=%s input_records=%s",
            self.db_path,
            self.table_name,
            account,
            folder_name,
            len(records),
        )

        try:
            async with aiosqlite.connect(self.db_path) as conn:
                await conn.execute("PRAGMA journal_mode=WAL")
                await conn.execute("PRAGMA synchronous=NORMAL")
                await self._ensure_schema(conn)

                written_count = 0
                for record in records:
                    message_id = str(record.get("message_id", "")).strip()
                    if not message_id:
                        continue
                    md5_value = self.message_id_md5(message_id)
                    if not md5_value:
                        continue
                    md5_by_message_id[message_id] = md5_value

                    base64a_value = str(record.get("Base64A", "") or "").strip()
                    if not base64a_value:
                        continue

                    sender_raw = str(record.get("sender", "") or "").strip()
                    sender_email = self.extract_sender_email(sender_raw)

                    await conn.execute(
                        f"""
                        INSERT INTO {self._table_sql} (
                            message_id_md5,
                            message_id,
                            account,
                            folder_name,
                            uid,
                            mail_id,
                            sender,
                            sender_email,
                            received_unixtime_ms,
                            base64a,
                            updated_unixtime_ms
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(message_id_md5) DO UPDATE SET
                            message_id = excluded.message_id,
                            account = excluded.account,
                            folder_name = excluded.folder_name,
                            uid = excluded.uid,
                            mail_id = excluded.mail_id,
                            sender = excluded.sender,
                            sender_email = excluded.sender_email,
                            received_unixtime_ms = excluded.received_unixtime_ms,
                            base64a = excluded.base64a,
                            updated_unixtime_ms = excluded.updated_unixtime_ms
                        """,
                        (
                            md5_value,
                            message_id,
                            account,
                            folder_name,
                            str(record.get("uid", "")).strip(),
                            str(record.get("mail_id", "")).strip(),
                            sender_raw,
                            sender_email,
                            self._safe_int(record.get("received_unixtime_ms", 0)),
                            base64a_value,
                            now_ms,
                        ),
                    )
                    written_count += 1
                await conn.commit()
                self.logger.info(
                    "sqlite upsert done: db=%s table=%s account=%s folder=%s mapped=%s written=%s",
                    self.db_path,
                    self.table_name,
                    account,
                    folder_name,
                    len(md5_by_message_id),
                    written_count,
                )
        except Exception:
            self.logger.exception(
                "sqlite upsert failed: db=%s table=%s account=%s folder=%s records=%s exists=%s size=%s",
                self.db_path,
                self.table_name,
                account,
                folder_name,
                len(records),
                self.db_path.exists(),
                self.db_path.stat().st_size if self.db_path.exists() else -1,
            )
            raise

        return md5_by_message_id

    async def _ensure_schema(self, conn: aiosqlite.Connection) -> None:
        await conn.execute(
            f"""
            CREATE TABLE IF NOT EXISTS {self._table_sql} (
                message_id_md5 TEXT PRIMARY KEY,
                message_id TEXT NOT NULL,
                account TEXT NOT NULL,
                folder_name TEXT NOT NULL,
                uid TEXT NOT NULL DEFAULT '',
                mail_id TEXT NOT NULL DEFAULT '',
                sender TEXT NOT NULL DEFAULT '',
                sender_email TEXT NOT NULL DEFAULT '',
                received_unixtime_ms INTEGER NOT NULL DEFAULT 0,
                base64a TEXT NOT NULL,
                updated_unixtime_ms INTEGER NOT NULL
            )
            """
        )
        await conn.execute(
            f"""
            CREATE INDEX IF NOT EXISTS {self._index_sql}
            ON {self._table_sql} (account, folder_name, uid)
            """
        )
        logging.getLogger(__name__).debug("sqlite schema ensured: table=%s", self.table_name)

    def fetch_by_md5(self, message_id_md5: str) -> dict | None:
        sql = f"""
        SELECT
            message_id_md5,
            message_id,
            account,
            folder_name,
            uid,
            mail_id,
            sender,
            sender_email,
            received_unixtime_ms,
            updated_unixtime_ms,
            base64a
        FROM {self._table_sql}
        WHERE message_id_md5 = ?
        LIMIT 1
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                row = conn.execute(sql, (message_id_md5,)).fetchone()
                if row is None:
                    self.logger.warning("sqlite fetch miss: db=%s table=%s md5=%s", self.db_path, self.table_name, message_id_md5)
                    return None
                self.logger.info("sqlite fetch hit: db=%s table=%s md5=%s", self.db_path, self.table_name, message_id_md5)
                return dict(row)
        except Exception:
            self.logger.exception(
                "sqlite fetch failed: db=%s table=%s md5=%s exists=%s size=%s",
                self.db_path,
                self.table_name,
                message_id_md5,
                self.db_path.exists(),
                self.db_path.stat().st_size if self.db_path.exists() else -1,
            )
            raise

    def fetch_by_message_id(self, message_id: str) -> dict | None:
        md5_value = self.message_id_md5(message_id)
        if not md5_value:
            return None
        return self.fetch_by_md5(md5_value)

    def table_exists(self) -> bool:
        if not self.db_path.exists():
            return False
        try:
            with sqlite3.connect(self.db_path) as conn:
                row = conn.execute(
                    "SELECT count(*) FROM sqlite_master WHERE type='table' AND name=?",
                    (self.table_name,),
                ).fetchone()
            return bool(row and int(row[0]) > 0)
        except Exception:
            self.logger.warning(
                "check sqlite table exists failed: db=%s table=%s",
                self.db_path,
                self.table_name,
                exc_info=True,
            )
            return False

    @staticmethod
    def extract_sender_email(sender_text: str) -> str:
        text = (sender_text or "").strip()
        if not text:
            return ""
        match = re.search(r"<([^<>\s@]+@[^<>\s@]+)>", text)
        if match:
            return match.group(1).strip().lower()
        match = re.search(r"\b([^<>\s@]+@[^<>\s@]+)\b", text)
        if match:
            return match.group(1).strip().lower()
        return ""

    @staticmethod
    def _safe_int(value: Any, default: int = 0) -> int:
        try:
            return int(value)
        except (TypeError, ValueError):
            return default
