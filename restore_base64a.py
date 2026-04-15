import argparse
import logging
import os
from pathlib import Path

import imap_outlook_oauth2
from public.email_ops import EmailOps
from public.file_ops import FileOps
from public.sqlite_ops import Base64ASqliteStore


def resolve_md5(message_id_md5: str, message_id: str) -> str:
    md5_value = (message_id_md5 or "").strip().lower()
    if md5_value:
        return md5_value
    return Base64ASqliteStore.message_id_md5((message_id or "").strip())


def resolve_provider_and_account() -> tuple[str, str]:
    profile = os.getenv("OUTLOOK_PROFILE", "outlook").strip() or "outlook"
    try:
        config_path = Path(imap_outlook_oauth2.resolve_default_config_path())
        row = imap_outlook_oauth2.OutlookMailService.load_outlook_config(config_path, profile)
        provider = str(row.get("mail", "")).strip() or profile
        account = str(row.get("user", "")).strip()
        return provider, account
    except Exception:
        return profile, ""


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Restore Base64A raw mail from provider db by message_id_md5 or message_id"
    )
    parser.add_argument("--db", default="", help="sqlite database path, default: db/<provider>.db")
    parser.add_argument("--table", default="", help="sqlite table name, default: base64A_<account>")
    parser.add_argument("--account", default="", help="account used to infer default table name")
    parser.add_argument("--md5", default="", help="message_id md5 (32 hex)")
    parser.add_argument("--message-id", default="", help="original message_id (will compute md5)")
    parser.add_argument("--save-eml", default="", help="output .eml file path")
    parser.add_argument("--save-base64", default="", help="output Base64A text file path")
    parser.add_argument("--save-html", default="", help="output html file path (extract from raw email)")
    parser.add_argument("--save-summary", default="", help="output decoded summary txt file path")
    parser.add_argument(
        "--print-preview",
        action="store_true",
        help="print first 500 chars of decoded text preview",
    )
    return parser


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s - %(message)s")
    logger = logging.getLogger("restore_base64a")

    args = build_parser().parse_args()
    provider, default_account = resolve_provider_and_account()
    safe_provider = Base64ASqliteStore.normalize_identifier(provider, fallback="outlook").lower()

    db_path = Path(args.db).resolve() if args.db else (Path("db") / f"{safe_provider}.db").resolve()
    account = (args.account or default_account or "mail").strip()
    table_name = args.table or Base64ASqliteStore.build_account_table_name(account)

    if not db_path.exists():
        print(f"[ERROR] DB not found: {db_path}")
        return 1

    md5_value = resolve_md5(args.md5, args.message_id)
    if not md5_value:
        print("[ERROR] Please provide --md5 or --message-id")
        return 1

    store = Base64ASqliteStore(db_path, table_name=table_name)
    try:
        record = store.fetch_by_md5(md5_value)
    except Exception as exc:
        logger.exception("query sqlite failed: db=%s table=%s md5=%s", db_path, table_name, md5_value)
        print(f"[ERROR] Query failed: {type(exc).__name__}: {exc}")
        return 4

    if record is None:
        print(f"[ERROR] Record not found: md5={md5_value} table={store.table_name}")
        return 2

    base64a_text = str(record.get("base64a") or "")
    if not base64a_text:
        print(f"[ERROR] Empty base64a payload: md5={md5_value}")
        return 3

    raw_bytes = EmailOps.decode_base64a_to_bytes(base64a_text)

    print("[OK] Record found")
    print(f"  db={db_path}")
    print(f"  table={store.table_name}")
    print(f"  md5={record['message_id_md5']}")
    print(f"  message_id={record['message_id']}")
    print(f"  account={record['account']}")
    print(f"  folder={record['folder_name']}")
    print(f"  uid={record['uid']}")
    print(f"  sender_email={record.get('sender_email', '')}")
    print(f"  base64_len={len(base64a_text)} decoded_bytes={len(raw_bytes)}")

    if args.save_base64:
        base64_path = FileOps.write_text(Path(args.save_base64), base64a_text, encoding="utf-8")
        print(f"  saved_base64={base64_path}")

    if args.save_eml:
        eml_path = FileOps.write_bytes(Path(args.save_eml), raw_bytes)
        print(f"  saved_eml={eml_path}")

    if args.save_html:
        html_text = EmailOps.extract_html(raw_bytes)
        html_path = FileOps.write_text(Path(args.save_html), html_text, encoding="utf-8")
        print(f"  saved_html={html_path}")
        print(f"  html_chars={len(html_text)}")
        if not html_text.strip():
            print("  [WARN] No text/html part found in this mail.")

    if args.save_summary:
        summary_text = EmailOps.build_readable_summary(raw_bytes, max_body_chars=0)
        summary_path = FileOps.write_text(Path(args.save_summary), summary_text, encoding="utf-8")
        print(f"  saved_summary={summary_path}")

    if args.print_preview:
        preview = raw_bytes.decode("utf-8", errors="replace")[:500]
        print("----- preview (first 500 chars) -----")
        print(preview)
        print("----- preview end -----")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
