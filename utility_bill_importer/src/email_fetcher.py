"""Fetch bill PDF attachments from an IMAP mailbox (e.g. Gmail)."""

from __future__ import annotations

import email
import imaplib
import re
from datetime import datetime, timedelta
from email.header import decode_header
from email.message import Message
from pathlib import Path


def _decode(value: str | bytes | None) -> str:
    if value is None:
        return ""
    parts = decode_header(value if isinstance(value, str) else value.decode("latin-1"))
    out = ""
    for text, enc in parts:
        if isinstance(text, bytes):
            out += text.decode(enc or "utf-8", errors="replace")
        else:
            out += text
    return out


def _safe_filename(name: str) -> str:
    name = _decode(name)
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("_")
    return name or "attachment.pdf"


def fetch_bill_pdfs(email_creds: dict, email_cfg: dict, dest_dir: Path) -> list[Path]:
    """Connect via IMAP, save matching attachments to dest_dir, return paths.

    Filtering is driven by config.yaml -> email:
      search_since_days, unseen_only, attachment_extensions, allowed_senders,
      mark_as_seen.
    """
    dest_dir.mkdir(parents=True, exist_ok=True)
    extensions = tuple(e.lower() for e in email_cfg.get("attachment_extensions", [".pdf"]))
    allowed_senders = [s.lower() for s in email_cfg.get("allowed_senders", []) if s]
    since_days = int(email_cfg.get("search_since_days", 45))
    unseen_only = bool(email_cfg.get("unseen_only", False))
    mark_seen = bool(email_cfg.get("mark_as_seen", False))

    if not email_creds.get("user") or not email_creds.get("password"):
        raise RuntimeError(
            "IMAP_USER / IMAP_PASSWORD not set. Copy .env.example to .env and fill it in."
        )

    saved: list[Path] = []
    conn = imaplib.IMAP4_SSL(email_creds["host"], email_creds["port"])
    try:
        conn.login(email_creds["user"], email_creds["password"])
        # Gmail labels with spaces must be quoted.
        status, _ = conn.select(f'"{email_creds["folder"]}"', readonly=not mark_seen)
        if status != "OK":
            raise RuntimeError(f"Cannot open mailbox/folder '{email_creds['folder']}'.")

        since = (datetime.now() - timedelta(days=since_days)).strftime("%d-%b-%Y")
        criteria = ["SINCE", since]
        if unseen_only:
            criteria = ["UNSEEN"] + criteria

        status, data = conn.search(None, *criteria)
        if status != "OK":
            return saved

        msg_ids = data[0].split()
        for mid in msg_ids:
            status, msg_data = conn.fetch(mid, "(RFC822)")
            if status != "OK" or not msg_data or not msg_data[0]:
                continue
            msg: Message = email.message_from_bytes(msg_data[0][1])

            sender = _decode(msg.get("From", "")).lower()
            if allowed_senders and not any(s in sender for s in allowed_senders):
                continue

            new_files = _save_attachments(msg, dest_dir, extensions)
            saved.extend(new_files)

            if mark_seen and new_files:
                conn.store(mid, "+FLAGS", "\\Seen")
    finally:
        try:
            conn.logout()
        except Exception:
            pass

    return saved


def _save_attachments(msg: Message, dest_dir: Path, extensions: tuple[str, ...]) -> list[Path]:
    saved: list[Path] = []
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        if part.get("Content-Disposition") is None:
            continue
        filename = part.get_filename()
        if not filename:
            continue
        safe = _safe_filename(filename)
        if not safe.lower().endswith(extensions):
            continue

        payload = part.get_payload(decode=True)
        if not payload:
            continue

        # Prefix with date to reduce collisions; skip if identical file exists.
        msg_date = msg.get("Date", "")
        prefix = re.sub(r"[^0-9]", "", msg_date)[:8] or datetime.now().strftime("%Y%m%d")
        out_path = dest_dir / f"{prefix}_{safe}"
        if out_path.exists() and out_path.stat().st_size == len(payload):
            continue
        out_path.write_bytes(payload)
        saved.append(out_path)
    return saved
