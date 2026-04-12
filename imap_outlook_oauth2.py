import argparse
import asyncio
import base64
import binascii
import csv
import email
import imaplib
import json
import logging
import os
import re
import sys
from dataclasses import dataclass
from email.header import decode_header
from html import unescape
from pathlib import Path

import aiohttp

try:
    import msal
except ImportError:
    msal = None


@dataclass
class OutlookConfig:
    email_addr: str
    client_id: str
    tenant: str
    host: str
    port: int
    mailbox: str
    scopes: list[str]
    cache_path: Path
    refresh_token: str
    csv_config_path: Path
    profile: str


class OutlookMailService:
    def __init__(self, config: OutlookConfig, logger: logging.Logger):
        self.config = config
        self.logger = logger

    @staticmethod
    def safe_base64_decode(data) -> bytes:
        _str_len = len(data)
        if _str_len % 4 != 0:
            return base64.urlsafe_b64decode(data + _str_len % 4 * b"=")
        return base64.urlsafe_b64decode(data)

    @classmethod
    def decode_csv_field(cls, value: str) -> str:
        raw = (value or "").strip()
        if not raw:
            return ""

        if not re.fullmatch(r"[A-Za-z0-9_-]+", raw):
            return raw

        try:
            decoded = cls.safe_base64_decode(raw.encode("utf-8"))
            text = decoded.decode("utf-8").strip()
            if not text:
                return raw
            # 只接受可打印文本，避免把普通明文误判为 Base64 后得到乱码。
            if not all(ch.isprintable() or ch in "\r\n\t" for ch in text):
                return raw
            return text
        except (binascii.Error, ValueError):
            return raw

    @classmethod
    def load_outlook_config(cls, csv_path: Path, profile: str = "outlook") -> dict[str, str]:
        if not csv_path.exists():
            return {}

        with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            rows = list(reader)

        if not rows:
            return {}

        selected = None
        for row in rows:
            if (row.get("mail", "") or "").strip().lower() == profile.lower():
                selected = row
                break
        if selected is None:
            selected = rows[0]

        return {
            "mail": (selected.get("mail", "") or "").strip(),
            "user": cls.decode_csv_field(selected.get("user", "")),
            "password": cls.decode_csv_field(selected.get("password", "")),
            "client_id": cls.decode_csv_field(selected.get("client_id", "")),
            "refresh_token": cls.decode_csv_field(selected.get("refresh_token", "")),
        }

    @staticmethod
    def decode_mime_words(value: str) -> str:
        if not value:
            return ""
        out = []
        for text, enc in decode_header(value):
            if isinstance(text, bytes):
                out.append(text.decode(enc or "utf-8", errors="replace"))
            else:
                out.append(text)
        return "".join(out)

    @staticmethod
    def build_xoauth2(user: str, access_token: str) -> bytes:
        return f"user={user}\x01auth=Bearer {access_token}\x01\x01".encode("utf-8")

    @staticmethod
    def parse_list_line(line: bytes) -> tuple[list[str], str]:
        raw = (line or b"").decode("utf-8", errors="replace").strip()
        # 典型格式: (\HasNoChildren \Junk) "/" Junk
        m = re.match(r"^\((?P<flags>[^)]*)\)\s+\"[^\"]*\"\s+(?P<name>.+)$", raw)
        if not m:
            return [], raw
        flags = [f for f in m.group("flags").split() if f]
        name = m.group("name").strip()
        if name.startswith('"') and name.endswith('"') and len(name) >= 2:
            name = name[1:-1]
        return flags, name

    def list_mailboxes(self, imap) -> list[dict[str, object]]:
        typ, data = imap.list()
        if typ != "OK" or data is None:
            return []
        result = []
        for line in data:
            flags, name = self.parse_list_line(line)
            result.append({"name": name, "flags": flags})
        return result

    def resolve_mailbox_name(self, imap, requested_mailbox: str) -> tuple[str, list[dict[str, object]]]:
        mailboxes = self.list_mailboxes(imap)
        requested = (requested_mailbox or "").strip()
        if not requested:
            return "INBOX", mailboxes

        for item in mailboxes:
            name = str(item.get("name", ""))
            if requested.lower() == name.lower():
                return name, mailboxes

        alias_to_flag = {
            "junk": r"\Junk",
            "junk email": r"\Junk",
            "spam": r"\Junk",
            "垃圾邮箱": r"\Junk",
            "垃圾邮件": r"\Junk",
            "trash": r"\Trash",
            "deleted": r"\Trash",
            "已删除": r"\Trash",
            "删除邮件": r"\Trash",
        }
        wanted_flag = alias_to_flag.get(requested.lower())
        if wanted_flag:
            for item in mailboxes:
                flags = [str(x) for x in item.get("flags", [])]
                if wanted_flag in flags:
                    return str(item.get("name", requested_mailbox)), mailboxes
        return requested_mailbox, mailboxes

    def get_public_client_app(self):
        token_cache = msal.SerializableTokenCache()
        if self.config.cache_path.exists():
            token_cache.deserialize(self.config.cache_path.read_text(encoding="utf-8"))
        authority = f"https://login.microsoftonline.com/{self.config.tenant}"
        app = msal.PublicClientApplication(
            client_id=self.config.client_id,
            authority=authority,
            token_cache=token_cache,
        )
        return app, token_cache

    def save_cache_if_changed(self, cache):
        if cache.has_state_changed:
            self.config.cache_path.write_text(cache.serialize(), encoding="utf-8")

    def acquire_access_token(self) -> str:
        app, cache = self.get_public_client_app()

        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(scopes=self.config.scopes, account=accounts[0])
            if result and "access_token" in result:
                self.save_cache_if_changed(cache)
                self.logger.info("通过缓存获取 access_token")
                return result["access_token"]

        flow = app.initiate_device_flow(scopes=self.config.scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"无法启动设备码登录: {flow}")

        print(flow.get("message", "请按设备码流程完成登录"))
        result = app.acquire_token_by_device_flow(flow)
        self.save_cache_if_changed(cache)

        if "access_token" not in result:
            raise RuntimeError(f"获取 access_token 失败: {json.dumps(result, ensure_ascii=False)}")

        self.logger.info("通过设备码流程获取 access_token")
        return result["access_token"]

    def acquire_access_token_by_refresh_token(self) -> str:
        async def _request_token() -> str:
            token_url = f"https://login.microsoftonline.com/{self.config.tenant}/oauth2/v2.0/token"
            payload = {
                "client_id": self.config.client_id,
                "grant_type": "refresh_token",
                "refresh_token": self.config.refresh_token,
                "scope": " ".join(self.config.scopes),
            }
            timeout = aiohttp.ClientTimeout(total=30)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.post(token_url, data=payload) as response:
                    text = await response.text()
                    if response.status != 200:
                        raise RuntimeError(
                            f"refresh_token 换取 access_token 失败: HTTP {response.status}, body={text}"
                        )
                    data = json.loads(text)
                    access_token = data.get("access_token")
                    if not access_token:
                        raise RuntimeError(f"返回中没有 access_token: {json.dumps(data, ensure_ascii=False)}")
                    return access_token

        token = asyncio.run(_request_token())
        self.logger.info("通过 refresh_token 获取 access_token")
        return token

    @staticmethod
    def html_to_text(html_content: str) -> str:
        text = re.sub(r"(?is)<(script|style).*?>.*?</\\1>", "", html_content)
        text = re.sub(r"(?s)<[^>]+>", " ", text)
        text = unescape(text)
        text = re.sub(r"[ \t]+", " ", text)
        text = re.sub(r"\n\s*\n+", "\n\n", text)
        return text.strip()

    @classmethod
    def extract_body_text(cls, msg: email.message.Message) -> str:
        if msg.is_multipart():
            html_candidate = ""
            for part in msg.walk():
                content_type = part.get_content_type()
                disposition = str(part.get("Content-Disposition", "")).lower()
                if "attachment" in disposition:
                    continue

                payload = part.get_payload(decode=True) or b""
                charset = part.get_content_charset() or "utf-8"
                text = payload.decode(charset, errors="replace")
                if content_type == "text/plain":
                    return text.strip()
                if content_type == "text/html" and not html_candidate:
                    html_candidate = text
            return cls.html_to_text(html_candidate) if html_candidate else ""

        payload = msg.get_payload(decode=True) or b""
        charset = msg.get_content_charset() or "utf-8"
        text = payload.decode(charset, errors="replace")
        if msg.get_content_type() == "text/html":
            return cls.html_to_text(text)
        return text.strip()

    def fetch_all_mails(self):
        self.logger.info("开始连接 IMAP: %s:%s", self.config.host, self.config.port)
        imap = imaplib.IMAP4_SSL(self.config.host, self.config.port)
        try:
            token = (
                self.acquire_access_token_by_refresh_token()
                if self.config.refresh_token
                else self.acquire_access_token()
            )

            xoauth2 = self.build_xoauth2(self.config.email_addr, token)
            imap.authenticate("XOAUTH2", lambda _: xoauth2)
            self.logger.info("IMAP 认证成功")

            mailbox_name, known_mailboxes = self.resolve_mailbox_name(imap, self.config.mailbox)
            typ, _ = imap.select(mailbox_name)
            if typ != "OK":
                names = [str(x.get("name", "")) for x in known_mailboxes]
                raise RuntimeError(
                    f"无法选择邮箱目录: {self.config.mailbox} (解析后: {mailbox_name})。可用目录: {', '.join(names)}"
                )

            typ, data = imap.search(None, "ALL")
            if typ != "OK":
                raise RuntimeError("搜索邮件失败")

            if not data or not data[0]:
                self.logger.info("邮箱无邮件")
                return {"total": 0, "mails": []}

            msg_ids = data[0].split()
            mails = []
            for idx, msg_id in enumerate(msg_ids, start=1):
                typ, msg_data = imap.fetch(msg_id, "(BODY.PEEK[])")
                if typ != "OK" or not msg_data or not msg_data[0] or not msg_data[0][1]:
                    subject = "(读取失败)"
                    body = ""
                else:
                    raw_message = msg_data[0][1] or b""
                    msg = email.message_from_bytes(raw_message)
                    subject = self.decode_mime_words(msg.get("Subject", "")) or "(无主题)"
                    body = self.extract_body_text(msg) or "(无可读正文)"

                mails.append(
                    {
                        "id": msg_id.decode("ascii", errors="ignore"),
                        "subject": subject,
                        "body": body,
                    }
                )
                self.logger.debug("已读取邮件 %s/%s: id=%s", idx, len(msg_ids), mails[-1]["id"])

            self.logger.info("读取完成，共 %s 封", len(msg_ids))
            return {"total": len(msg_ids), "mails": mails}
        finally:
            try:
                imap.logout()
                self.logger.info("IMAP 连接已关闭")
            except Exception:
                self.logger.debug("IMAP 登出时忽略异常", exc_info=True)

    def get_mailboxes(self):
        self.logger.info("开始连接 IMAP: %s:%s", self.config.host, self.config.port)
        imap = imaplib.IMAP4_SSL(self.config.host, self.config.port)
        try:
            token = (
                self.acquire_access_token_by_refresh_token()
                if self.config.refresh_token
                else self.acquire_access_token()
            )
            xoauth2 = self.build_xoauth2(self.config.email_addr, token)
            imap.authenticate("XOAUTH2", lambda _: xoauth2)
            self.logger.info("IMAP 认证成功")
            return self.list_mailboxes(imap)
        finally:
            try:
                imap.logout()
                self.logger.info("IMAP 连接已关闭")
            except Exception:
                self.logger.debug("IMAP 登出时忽略异常", exc_info=True)


def build_runtime_config(args) -> OutlookConfig:
    csv_config_path = Path(args.config)
    csv_config = OutlookMailService.load_outlook_config(csv_config_path, args.profile)

    email_addr = os.getenv("OUTLOOK_EMAIL", csv_config.get("user", ""))
    client_id = os.getenv("OUTLOOK_CLIENT_ID", csv_config.get("client_id", ""))
    if not email_addr:
        raise RuntimeError("缺少邮箱配置: OUTLOOK_EMAIL 或 config/OutLook.csv 的 user")
    if not client_id:
        raise RuntimeError("缺少 client_id 配置: OUTLOOK_CLIENT_ID 或 config/OutLook.csv 的 client_id")

    return OutlookConfig(
        email_addr=email_addr,
        client_id=client_id,
        tenant=os.getenv("OUTLOOK_TENANT", "consumers"),
        host=os.getenv("OUTLOOK_IMAP_HOST", "outlook.office365.com"),
        port=int(os.getenv("OUTLOOK_IMAP_PORT", "993")),
        mailbox=args.mailbox,
        scopes=[
            "https://outlook.office.com/IMAP.AccessAsUser.All",
            "offline_access",
        ],
        cache_path=Path(os.getenv("OUTLOOK_TOKEN_CACHE", ".outlook_token_cache.json")),
        refresh_token=os.getenv("OUTLOOK_REFRESH_TOKEN", csv_config.get("refresh_token", "")),
        csv_config_path=csv_config_path,
        profile=args.profile,
    )


def resolve_default_config_path() -> str:
    config_from_env = os.getenv("OUTLOOK_CONFIG_PATH", "").strip()
    if config_from_env:
        return config_from_env

    local_config = Path("config/OutLook.local.csv")
    if local_config.exists():
        return str(local_config)
    return "config/OutLook.csv"


def mask_email(value: str) -> str:
    if not value or "@" not in value:
        return "***"
    local, _, domain = value.partition("@")
    if len(local) <= 2:
        safe_local = local[:1] + "*"
    else:
        safe_local = local[:2] + "*" * (len(local) - 2)
    return f"{safe_local}@{domain}"


def mask_secret(value: str, keep_start: int = 4, keep_end: int = 4) -> str:
    if not value:
        return "***"
    if len(value) <= keep_start + keep_end:
        return "*" * len(value)
    return f"{value[:keep_start]}***{value[-keep_end:]}"


def setup_logger(level: str, log_file: str) -> logging.Logger:
    logger = logging.getLogger("outlook_mail")
    logger.handlers.clear()
    logger.setLevel(getattr(logging, level.upper(), logging.INFO))

    fmt = logging.Formatter("%(asctime)s %(levelname)s %(name)s - %(message)s")

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(fmt)
    logger.addHandler(stream_handler)

    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(log_path, encoding="utf-8")
        file_handler.setFormatter(fmt)
        logger.addHandler(file_handler)

    return logger


def main() -> int:
    parser = argparse.ArgumentParser(description="Outlook IMAP OAuth2 收取未读邮件示例")
    parser.add_argument("--dry-run", action="store_true", help="只检查配置，不连接网络")
    parser.add_argument("--list-mailboxes", action="store_true", help="列出邮箱目录，不读取邮件内容")
    parser.add_argument("--mailbox", default=os.getenv("OUTLOOK_IMAP_MAILBOX", "INBOX"))
    parser.add_argument("--profile", default=os.getenv("OUTLOOK_PROFILE", "outlook"), help="CSV 配置中的 mail 字段")
    parser.add_argument(
        "--config",
        default=resolve_default_config_path(),
        help="Outlook CSV 配置文件路径",
    )
    parser.add_argument("--log-level", default=os.getenv("OUTLOOK_LOG_LEVEL", "INFO"), help="日志级别")
    parser.add_argument("--log-file", default=os.getenv("OUTLOOK_LOG_FILE", ""), help="日志文件路径")
    args = parser.parse_args()

    logger = setup_logger(args.log_level, args.log_file)

    if msal is None:
        print("缺少依赖 msal/aiohttp，请先执行: D:\\0Code2\\py312\\python.exe -m pip install msal aiohttp", file=sys.stderr)
        return 2

    config = build_runtime_config(args)
    service = OutlookMailService(config, logger)

    if args.dry_run:
        print("配置检查通过:")
        print(f"  email={mask_email(config.email_addr)}")
        print(f"  client_id={mask_secret(config.client_id)}")
        print(f"  tenant={config.tenant}")
        print(f"  host={config.host}:{config.port}")
        print(f"  mailbox={config.mailbox}")
        print(f"  cache={config.cache_path.resolve()}")
        print(f"  config={config.csv_config_path.resolve()}")
        print(f"  profile={config.profile}")
        print(f"  auth_mode={'refresh_token' if config.refresh_token else 'device_flow_or_cache'}")
        logger.info("dry-run 完成")
        return 0

    if args.list_mailboxes:
        boxes = service.get_mailboxes()
        print("邮箱目录列表:")
        for item in boxes:
            name = item.get("name", "")
            flags = " ".join(item.get("flags", []))
            print(f"- {name}  [{flags}]")
        return 0

    result = service.fetch_all_mails()
    print(f"邮件总数: {result['total']}")
    if result["total"] == 0:
        return 0

    print("邮件内容列表:")
    for item in result["mails"]:
        print(f"\nID: {item['id']}")
        print(f"标题: {item['subject']}")
        print("正文:")
        print(item["body"])
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("键盘退出")
        sys.exit(0)
    except Exception as exc:
        print(f"运行失败: {exc}", file=sys.stderr)
        raise SystemExit(1)
