import base64
import logging
from email import policy
from email.header import decode_header, make_header
from email.parser import BytesParser


class EmailOps:
    logger = logging.getLogger(__name__)

    @staticmethod
    def add_padding(urlsafe_b64: str) -> str:
        missing = (-len(urlsafe_b64)) % 4
        if missing == 0:
            return urlsafe_b64
        return f"{urlsafe_b64}{'=' * missing}"

    @staticmethod
    def decode_base64a_to_bytes(base64a_text: str) -> bytes:
        padded = EmailOps.add_padding(base64a_text)
        data = base64.urlsafe_b64decode(padded.encode("utf-8"))
        EmailOps.logger.info("decode base64a: input_chars=%s output_bytes=%s", len(base64a_text), len(data))
        return data

    @staticmethod
    def decode_mime_header(value: str) -> str:
        text = (value or "").strip()
        if not text:
            return ""
        try:
            return str(make_header(decode_header(text)))
        except Exception:
            return text

    @staticmethod
    def _decode_part_payload(part) -> str:
        payload = part.get_payload(decode=True)
        if payload is None:
            return ""
        charset = part.get_content_charset() or "utf-8"
        try:
            return payload.decode(charset, errors="replace")
        except LookupError:
            return payload.decode("utf-8", errors="replace")

    @staticmethod
    def extract_html(raw_bytes: bytes) -> str:
        msg = BytesParser(policy=policy.default).parsebytes(raw_bytes)
        html_parts: list[str] = []
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    html_text = EmailOps._decode_part_payload(part)
                    if html_text.strip():
                        html_parts.append(html_text)
        elif msg.get_content_type() == "text/html":
            html_text = EmailOps._decode_part_payload(msg)
            if html_text.strip():
                html_parts.append(html_text)
        html = "\n\n<!-- part split -->\n\n".join(html_parts)
        EmailOps.logger.info("extract html: parts=%s chars=%s", len(html_parts), len(html))
        return html

    @staticmethod
    def extract_text_plain(raw_bytes: bytes) -> str:
        msg = BytesParser(policy=policy.default).parsebytes(raw_bytes)
        text_parts: list[str] = []
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    text = EmailOps._decode_part_payload(part)
                    if text.strip():
                        text_parts.append(text)
        elif msg.get_content_type() == "text/plain":
            text = EmailOps._decode_part_payload(msg)
            if text.strip():
                text_parts.append(text)
        text = "\n\n----- part split -----\n\n".join(text_parts)
        EmailOps.logger.info("extract text/plain: parts=%s chars=%s", len(text_parts), len(text))
        return text

    @staticmethod
    def build_readable_summary(raw_bytes: bytes, max_body_chars: int = 0) -> str:
        msg = BytesParser(policy=policy.default).parsebytes(raw_bytes)
        subject = EmailOps.decode_mime_header(str(msg.get("Subject", "")))
        sender = EmailOps.decode_mime_header(str(msg.get("From", "")))
        to = EmailOps.decode_mime_header(str(msg.get("To", "")))
        date = EmailOps.decode_mime_header(str(msg.get("Date", "")))
        message_id = EmailOps.decode_mime_header(str(msg.get("Message-ID", "")))

        text_plain = EmailOps.extract_text_plain(raw_bytes).strip()
        html_text = EmailOps.extract_html(raw_bytes).strip()
        body_source = "text/plain" if text_plain else "text/html"
        body_text = text_plain if text_plain else html_text

        if max_body_chars > 0 and len(body_text) > max_body_chars:
            body_preview = f"{body_text[:max_body_chars]}..."
        else:
            body_preview = body_text

        lines = [
            "Email Summary (decoded)",
            f"Subject: {subject}",
            f"From: {sender}",
            f"To: {to}",
            f"Date: {date}",
            f"Message-ID: {message_id}",
            f"Body-Source: {body_source}",
            "",
            "Body Preview:",
            body_preview,
            "",
        ]
        summary = "\n".join(lines)
        EmailOps.logger.info(
            "build summary: body_source=%s max_body_chars=%s summary_chars=%s",
            body_source,
            max_body_chars,
            len(summary),
        )
        return summary
