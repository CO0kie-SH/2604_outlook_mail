import argparse
import email
import imaplib
import logging
import select
import sys
import time
from pathlib import Path
from types import SimpleNamespace
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import imap_outlook_oauth2


def setup_logger(project_root: Path) -> logging.Logger:
    log_dir = project_root / "log"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{time.strftime('%Y%m%d')}_idle_validate.log"

    logger = logging.getLogger("validate_imap_idle_once")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s %(levelname)s %(name)s - %(message)s")
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(formatter)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(formatter)
    logger.addHandler(fh)
    logger.addHandler(sh)
    return logger


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run one-shot IMAP IDLE validation with existing Outlook config")
    parser.add_argument("--profile", default="outlook", help="config profile (mail field in OutLook.csv)")
    parser.add_argument("--config", default=imap_outlook_oauth2.resolve_default_config_path(), help="outlook config csv path")
    parser.add_argument("--mailbox", default="INBOX", help="mailbox name")
    parser.add_argument("--idle-seconds", type=int, default=90, help="IDLE wait seconds")
    parser.add_argument("--fetch-limit", type=int, default=5, help="max fetched new mails after IDLE event")
    return parser.parse_args()


def build_service(args: argparse.Namespace, logger: logging.Logger) -> imap_outlook_oauth2.OutlookMailService:
    ns = SimpleNamespace(
        config=args.config,
        profile=args.profile,
        mailbox=args.mailbox,
    )
    config = imap_outlook_oauth2.build_runtime_config(ns)
    return imap_outlook_oauth2.OutlookMailService(config, logger)


def acquire_token(service: imap_outlook_oauth2.OutlookMailService) -> str:
    if service.config.refresh_token:
        return service.acquire_access_token_by_refresh_token()
    return service.acquire_access_token()


def select_mailbox(service: imap_outlook_oauth2.OutlookMailService, imap, requested: str) -> str:
    mailbox_name, _ = service.resolve_mailbox_name(imap, requested)
    typ, _ = imap.select(mailbox_name)
    if typ != "OK":
        raise RuntimeError(f"select mailbox failed: {requested} -> {mailbox_name}")
    return mailbox_name


def fetch_uid_list(imap) -> list[int]:
    typ, data = imap.uid("SEARCH", None, "ALL")
    if typ != "OK" or not data or not data[0]:
        return []
    values: list[int] = []
    for token in data[0].split():
        try:
            values.append(int(token.decode("ascii", errors="ignore")))
        except Exception:
            continue
    return values


def _safe_readline(imap, timeout_seconds: float) -> bytes | None:
    sock = getattr(imap, "sock", None)
    if sock is None:
        raise RuntimeError("imap socket not available")
    readable, _, _ = select.select([sock], [], [], max(0.0, timeout_seconds))
    if not readable:
        return None
    line = imap.readline()
    if not isinstance(line, bytes):
        return None
    return line


def idle_wait_once(imap, idle_seconds: int, logger: logging.Logger) -> list[str]:
    if not hasattr(imap, "_new_tag"):
        raise RuntimeError("imaplib has no _new_tag, cannot run raw IDLE command")

    tag = imap._new_tag()  # type: ignore[attr-defined]
    tag_text = tag.decode("ascii", errors="ignore") if isinstance(tag, bytes) else str(tag)
    imap.send(f"{tag_text} IDLE\r\n".encode("ascii"))

    ack_line = _safe_readline(imap, timeout_seconds=10)
    if ack_line is None:
        raise RuntimeError("IDLE ack timeout")
    ack_text = ack_line.decode("utf-8", errors="replace").strip()
    if not ack_text.startswith("+"):
        raise RuntimeError(f"IDLE ack invalid: {ack_text}")
    logger.info("IDLE started: ack=%s", ack_text)

    events: list[str] = []
    deadline = time.time() + max(1, idle_seconds)
    triggered = False
    while time.time() < deadline:
        line = _safe_readline(imap, timeout_seconds=1.0)
        if line is None:
            continue
        text = line.decode("utf-8", errors="replace").strip()
        if not text:
            continue
        events.append(text)
        logger.info("IDLE event: %s", text)
        if " EXISTS" in text or " RECENT" in text or " EXPUNGE" in text:
            triggered = True
            break

    imap.send(b"DONE\r\n")
    done_deadline = time.time() + 10
    while time.time() < done_deadline:
        line = _safe_readline(imap, timeout_seconds=1.0)
        if line is None:
            continue
        text = line.decode("utf-8", errors="replace").strip()
        if not text:
            continue
        if text.startswith(tag_text):
            logger.info("IDLE finished: %s", text)
            break
        events.append(text)

    if not triggered:
        logger.info("IDLE finished without push event within %ss", idle_seconds)
    return events


def fetch_new_mail_headers(
    service: imap_outlook_oauth2.OutlookMailService,
    imap,
    baseline_max_uid: int,
    fetch_limit: int,
) -> list[dict[str, Any]]:
    all_uids = fetch_uid_list(imap)
    new_uids = [uid for uid in all_uids if uid > baseline_max_uid]
    if fetch_limit > 0 and len(new_uids) > fetch_limit:
        new_uids = new_uids[-fetch_limit:]

    rows: list[dict[str, Any]] = []
    for uid in new_uids:
        typ, msg_data = imap.uid("FETCH", str(uid), "(UID BODY.PEEK[HEADER.FIELDS (DATE SUBJECT FROM MESSAGE-ID)])")
        if typ != "OK" or not msg_data or not msg_data[0] or not msg_data[0][1]:
            continue
        msg = email.message_from_bytes(msg_data[0][1] or b"")
        rows.append(
            {
                "uid": str(uid),
                "message_id": str(msg.get("Message-ID", "")).strip(),
                "sender": service.decode_mime_words(msg.get("From", "")),
                "title": service.decode_mime_words(msg.get("Subject", "")),
                "received_at": service.parse_message_date(str(msg.get("Date", ""))),
            }
        )
    return rows


def main() -> int:
    args = parse_args()
    project_root = Path(__file__).resolve().parents[1]
    logger = setup_logger(project_root)
    logger.info("idle validation started: mailbox=%s idle_seconds=%s", args.mailbox, args.idle_seconds)

    service = build_service(args, logger)
    token = acquire_token(service)
    logger.info("token acquired, connecting IMAP host=%s port=%s", service.config.host, service.config.port)

    imap = imaplib.IMAP4_SSL(service.config.host, service.config.port)
    try:
        xoauth2 = service.build_xoauth2(service.config.email_addr, token)
        imap.authenticate("XOAUTH2", lambda _: xoauth2)
        mailbox_name = select_mailbox(service, imap, args.mailbox)
        logger.info("imap authenticated and mailbox selected: %s", mailbox_name)

        baseline_uids = fetch_uid_list(imap)
        baseline_count = len(baseline_uids)
        baseline_max_uid = max(baseline_uids) if baseline_uids else 0
        logger.info("baseline snapshot: count=%s max_uid=%s", baseline_count, baseline_max_uid)
        logger.info("now enter IDLE, send one new mail to this mailbox to trigger push")

        events = idle_wait_once(imap, idle_seconds=args.idle_seconds, logger=logger)
        latest_uids = fetch_uid_list(imap)
        latest_count = len(latest_uids)
        latest_max_uid = max(latest_uids) if latest_uids else 0
        logger.info("after IDLE snapshot: count=%s max_uid=%s events=%s", latest_count, latest_max_uid, len(events))

        new_headers = fetch_new_mail_headers(
            service=service,
            imap=imap,
            baseline_max_uid=baseline_max_uid,
            fetch_limit=max(1, int(args.fetch_limit)),
        )
        for idx, row in enumerate(new_headers, start=1):
            logger.info(
                "new_mail[%s] uid=%s received_at=%s sender=%s title=%s message_id=%s",
                idx,
                row.get("uid"),
                row.get("received_at"),
                row.get("sender"),
                row.get("title"),
                row.get("message_id"),
            )

        if latest_count > baseline_count or new_headers:
            logger.info("IDLE validation PASS: push observed and incremental fetch ran")
            return 0

        logger.warning("IDLE validation finished without observable new mail in this round")
        return 0
    finally:
        try:
            imap.logout()
        except Exception:
            logger.warning("imap logout failed", exc_info=True)


if __name__ == "__main__":
    raise SystemExit(main())
