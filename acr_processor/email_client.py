import email
import imaplib
import os
import re
from datetime import datetime, timedelta
from email.header import decode_header
from pathlib import Path
from typing import List


def _decode_mime_words(value: str) -> str:
    if not value:
        return ""

    parts = decode_header(value)
    decoded = []
    for text, encoding in parts:
        if isinstance(text, bytes):
            decoded.append(text.decode(encoding or "utf-8", errors="replace"))
        else:
            decoded.append(text)
    return "".join(decoded)


def _sanitize_filename(name: str) -> str:
    return re.sub(r"[^a-zA-Z0-9._-]", "_", name)


def _looks_like_acr_email(subject: str, body: str, keywords: List[str]) -> bool:
    haystack = f"{subject}\n{body}".lower()
    return any(keyword.lower().strip() in haystack for keyword in keywords if keyword.strip())


def _extract_plain_text_body(message: email.message.Message) -> str:
    if message.is_multipart():
        for part in message.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition", ""))
            if content_type == "text/plain" and "attachment" not in content_disposition.lower():
                payload = part.get_payload(decode=True)
                if payload:
                    return payload.decode(part.get_content_charset() or "utf-8", errors="replace")
    else:
        payload = message.get_payload(decode=True)
        if payload:
            return payload.decode(message.get_content_charset() or "utf-8", errors="replace")

    return ""


def _to_bool(value: str, default: bool = False) -> bool:
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def download_acr_pdfs(download_dir: Path, keywords: List[str]) -> List[dict]:
    host = os.getenv("IMAP_HOST")
    port = int(os.getenv("IMAP_PORT", "993"))
    username = os.getenv("IMAP_USERNAME")
    password = os.getenv("IMAP_PASSWORD")
    mailbox = os.getenv("IMAP_MAILBOX", "INBOX")
    days_back = int(os.getenv("IMAP_DAYS_BACK", "14"))
    max_emails = int(os.getenv("IMAP_MAX_EMAILS", "200"))
    unseen_only = _to_bool(os.getenv("IMAP_UNSEEN_ONLY", "true"), default=True)

    if not all([host, username, password]):
        raise ValueError("Missing IMAP settings. Check your .env file.")

    download_dir.mkdir(parents=True, exist_ok=True)
    saved_files = []

    with imaplib.IMAP4_SSL(host, port) as mail:
        mail.login(username, password)
        mail.select(mailbox)

        # Limit to recent/unread emails first to keep runs quick on large inboxes.
        search_terms = []
        if unseen_only:
            search_terms.append("UNSEEN")
        if days_back > 0:
            since_date = (datetime.now() - timedelta(days=days_back)).strftime("%d-%b-%Y")
            search_terms.extend(["SINCE", since_date])
        if not search_terms:
            search_terms = ["ALL"]

        status, data = mail.search(None, *search_terms)

        if status != "OK":
            return saved_files

        email_ids = data[0].split()
        email_ids = list(reversed(email_ids))
        if max_emails > 0:
            email_ids = email_ids[:max_emails]

        for email_id in email_ids:
            status, msg_data = mail.fetch(email_id, "(RFC822)")
            if status != "OK" or not msg_data:
                continue

            raw_email = msg_data[0][1]
            message = email.message_from_bytes(raw_email)

            subject = _decode_mime_words(message.get("Subject", ""))
            sender = _decode_mime_words(message.get("From", ""))
            body = _extract_plain_text_body(message)

            if not _looks_like_acr_email(subject, body, keywords):
                continue

            for part in message.walk():
                content_disposition = str(part.get("Content-Disposition", ""))
                filename = part.get_filename()
                if not filename:
                    continue

                decoded_filename = _decode_mime_words(filename)
                if "attachment" not in content_disposition.lower():
                    continue
                if not decoded_filename.lower().endswith(".pdf"):
                    continue

                safe_name = _sanitize_filename(decoded_filename)
                target_path = download_dir / f"{email_id.decode()}_{safe_name}"

                payload = part.get_payload(decode=True)
                if not payload:
                    continue

                with open(target_path, "wb") as f:
                    f.write(payload)

                saved_files.append(
                    {
                        "pdf_path": str(target_path),
                        "pdf_filename": safe_name,
                        "email_subject": subject,
                        "email_sender": sender,
                    }
                )

    return saved_files
