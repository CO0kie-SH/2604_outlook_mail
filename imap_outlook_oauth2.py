import argparse
import asyncio
import email
import imaplib
import json
import os
import re
import sys
from email.header import decode_header
from html import unescape
from pathlib import Path
import aiohttp

try:
    import msal
except ImportError:
    msal = None


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


def build_xoauth2(user: str, access_token: str) -> bytes:
    return f"user={user}\x01auth=Bearer {access_token}\x01\x01".encode("utf-8")


def get_public_client_app(client_id: str, tenant: str, cache_path: Path):
    token_cache = msal.SerializableTokenCache()
    if cache_path.exists():
        token_cache.deserialize(cache_path.read_text(encoding="utf-8"))
    authority = f"https://login.microsoftonline.com/{tenant}"
    app = msal.PublicClientApplication(client_id=client_id, authority=authority, token_cache=token_cache)
    return app, token_cache


def save_cache_if_changed(cache, cache_path: Path):
    if cache.has_state_changed:
        cache_path.write_text(cache.serialize(), encoding="utf-8")


def acquire_access_token(client_id: str, tenant: str, scopes: list[str], cache_path: Path) -> str:
    app, cache = get_public_client_app(client_id, tenant, cache_path)

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes=scopes, account=accounts[0])
        if result and "access_token" in result:
            save_cache_if_changed(cache, cache_path)
            return result["access_token"]

    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        raise RuntimeError(f"无法启动设备码登录: {flow}")

    print(flow.get("message", "请按设备码流程完成登录"))
    result = app.acquire_token_by_device_flow(flow)
    save_cache_if_changed(cache, cache_path)

    if "access_token" not in result:
        raise RuntimeError(f"获取 access_token 失败: {json.dumps(result, ensure_ascii=False)}")

    return result["access_token"]


def acquire_access_token_by_refresh_token(
    client_id: str,
    tenant: str,
    refresh_token: str,
    scopes: list[str],
) -> str:
    async def _request_token() -> str:
        token_url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
        payload = {
            "client_id": client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": " ".join(scopes),
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

    return asyncio.run(_request_token())


def html_to_text(html_content: str) -> str:
    text = re.sub(r"(?is)<(script|style).*?>.*?</\\1>", "", html_content)
    text = re.sub(r"(?s)<[^>]+>", " ", text)
    text = unescape(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n\n", text)
    return text.strip()


def extract_body_text(msg: email.message.Message) -> str:
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
        return html_to_text(html_candidate) if html_candidate else ""

    payload = msg.get_payload(decode=True) or b""
    charset = msg.get_content_charset() or "utf-8"
    text = payload.decode(charset, errors="replace")
    if msg.get_content_type() == "text/html":
        return html_to_text(text)
    return text.strip()


def fetch_all_mails(email_addr: str, access_token: str, host: str, port: int, mailbox: str):
    imap = imaplib.IMAP4_SSL(host, port)
    try:
        xoauth2 = build_xoauth2(email_addr, access_token)
        imap.authenticate("XOAUTH2", lambda _: xoauth2)

        typ, _ = imap.select(mailbox)
        if typ != "OK":
            raise RuntimeError(f"无法选择邮箱目录: {mailbox}")

        typ, data = imap.search(None, "ALL")
        if typ != "OK":
            raise RuntimeError("搜索邮件失败")

        if not data or not data[0]:
            return {"total": 0, "mails": []}

        msg_ids = data[0].split()
        mails = []
        for msg_id in msg_ids:
            typ, msg_data = imap.fetch(msg_id, "(BODY.PEEK[])")
            if typ != "OK" or not msg_data or not msg_data[0] or not msg_data[0][1]:
                subject = "(读取失败)"
                body = ""
            else:
                raw_message = msg_data[0][1] or b""
                msg = email.message_from_bytes(raw_message)
                subject = decode_mime_words(msg.get("Subject", "")) or "(无主题)"
                body = extract_body_text(msg) or "(无可读正文)"
            mails.append(
                {
                    "id": msg_id.decode("ascii", errors="ignore"),
                    "subject": subject,
                    "body": body,
                }
            )

        return {"total": len(msg_ids), "mails": mails}
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def required_env(name: str, default: str | None = None) -> str:
    value = os.getenv(name, default)
    if not value:
        raise RuntimeError(f"缺少环境变量: {name}")
    return value


def main() -> int:
    parser = argparse.ArgumentParser(description="Outlook IMAP OAuth2 收取未读邮件示例")
    parser.add_argument("--dry-run", action="store_true", help="只检查配置，不连接网络")
    parser.add_argument("--mailbox", default=os.getenv("OUTLOOK_IMAP_MAILBOX", "INBOX"))
    args = parser.parse_args()

    if msal is None:
        print("缺少依赖 msal/aiohttp，请先执行: D:\\0Code2\\py312\\python.exe -m pip install msal aiohttp", file=sys.stderr)
        return 2

    email_addr = required_env("OUTLOOK_EMAIL")
    client_id = required_env("OUTLOOK_CLIENT_ID")
    tenant = os.getenv("OUTLOOK_TENANT", "consumers")
    host = os.getenv("OUTLOOK_IMAP_HOST", "outlook.office365.com")
    port = int(os.getenv("OUTLOOK_IMAP_PORT", "993"))
    scopes = [
        "https://outlook.office.com/IMAP.AccessAsUser.All",
        "offline_access",
    ]
    cache_path = Path(os.getenv("OUTLOOK_TOKEN_CACHE", ".outlook_token_cache.json"))
    refresh_token = os.getenv("OUTLOOK_REFRESH_TOKEN", "")

    if args.dry_run:
        print("配置检查通过:")
        print(f"  email={email_addr}")
        print(f"  client_id={client_id}")
        print(f"  tenant={tenant}")
        print(f"  host={host}:{port}")
        print(f"  mailbox={args.mailbox}")
        print(f"  cache={cache_path.resolve()}")
        print(f"  auth_mode={'refresh_token' if refresh_token else 'device_flow_or_cache'}")
        return 0

    if refresh_token:
        token = acquire_access_token_by_refresh_token(client_id, tenant, refresh_token, scopes)
    else:
        token = acquire_access_token(client_id, tenant, scopes, cache_path)
    result = fetch_all_mails(email_addr, token, host, port, args.mailbox)
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
    pass
