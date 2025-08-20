#!/usr/bin/env python3
"""Send or draft emails with attachments loaded from a directory.

Usage example:
    python scripts/actions/send_email_with_attachments.py data.json \
        --attachments data/attachments --draft --templates templates/word --attach-pdf Yes

Environment variables used when ``--send`` is specified:
    SMTP_SERVER, SMTP_PORT, SMTP_USERNAME, SMTP_PASSWORD
"""
from __future__ import annotations

import argparse
import importlib
import json
import logging
import mimetypes
import os
import re
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable, List, Dict, Any, Optional
import smtplib

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def load_records(path: Path, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
    """Load email records from a JSON or Excel file.

    The file must contain objects with at least ``to``, ``subject`` and
    ``body`` (or ``html``) fields. Excel support requires the ``openpyxl`` package.
    If sheet_name is provided, that sheet will be read.
    """
    ext = path.suffix.lower()
    if ext == ".json":
        with path.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
        if isinstance(data, dict):
            data = [data]
        normalized = []
        for r in data:
            normalized.append({str(k).strip().lower(): v for k, v in r.items()})
        return normalized

    if ext in {".xls", ".xlsx"}:
        try:
            openpyxl = importlib.import_module("openpyxl")
        except ModuleNotFoundError as exc:
            raise ModuleNotFoundError("openpyxl is required to read Excel files") from exc
        wb = openpyxl.load_workbook(path, data_only=True)
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                raise ValueError(f"Sheet '{sheet_name}' not found in {path} (available: {wb.sheetnames})")
            ws = wb[sheet_name]
        else:
            ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []
        headers = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
        records: List[Dict[str, Any]] = []
        for row in rows[1:]:
            record = {}
            for i, val in enumerate(row):
                header = headers[i] if i < len(headers) else f"col_{i}"
                record[header] = val
            records.append(record)
        return records

    raise ValueError(f"Unsupported file extension: {ext}")


def attach_files(message: EmailMessage, directory: Path, include_pdfs: bool = True) -> None:
    """Attach files from ``directory`` to ``message``.

    If include_pdfs is False, PDF files will be skipped.
    """
    if not directory or not directory.exists():
        logging.debug("Attachments directory %s does not exist — skipping attachments.", directory)
        return

    for file_path in sorted(directory.glob("*")):
        if not file_path.is_file():
            continue
        if not include_pdfs and file_path.suffix.lower() == ".pdf":
            logging.debug("Skipping PDF due to --attach-pdf=No: %s", file_path.name)
            continue
        ctype, encoding = mimetypes.guess_type(str(file_path))
        if ctype is None:
            maintype, subtype = "application", "octet-stream"
        else:
            maintype, subtype = ctype.split("/", 1)
        with file_path.open("rb") as fh:
            data = fh.read()
        message.add_attachment(data, maintype=maintype, subtype=subtype, filename=file_path.name)


def attach_word_templates(message: EmailMessage, templates_dir: Optional[Path]) -> None:
    """Attach Word templates (.docx, .doc) from templates_dir to message (if provided)."""
    if not templates_dir:
        return
    if not templates_dir.exists():
        logging.warning("Templates directory does not exist: %s", templates_dir)
        return
    for file_path in sorted(templates_dir.glob("*")):
        if not file_path.is_file():
            continue
        if file_path.suffix.lower() not in {".docx", ".doc"}:
            continue
        with file_path.open("rb") as fh:
            data = fh.read()
        # use generic application/msword or appropriate mimetype
        ctype, _ = mimetypes.guess_type(str(file_path))
        if not ctype:
            maintype, subtype = "application", "octet-stream"
        else:
            maintype, subtype = ctype.split("/", 1)
        message.add_attachment(data, maintype=maintype, subtype=subtype, filename=file_path.name)


def sanitize_filename(s: str, max_len: int = 200) -> str:
    """Remove problematic chars and limit length for file names."""
    s = re.sub(r"[\\/:\*\?\"<>\|]+", "-", s or "")
    s = re.sub(r"\s+", " ", s).strip()
    if len(s) > max_len:
        s = s[:max_len]
    return s


def create_message(record: Dict[str, Any], default_attachments: Path, include_pdfs: bool, templates_dir: Optional[Path]) -> EmailMessage:
    msg = EmailMessage()
    to_field = record.get("to") or record.get("recipient") or ""
    subject = str(record.get("subject") or "")
    from_addr = os.environ.get("SMTP_USERNAME", "noreply@example.com")

    msg["To"] = to_field
    if record.get("cc"):
        msg["Cc"] = str(record.get("cc"))
    if record.get("bcc"):
        msg["Bcc"] = str(record.get("bcc"))
    msg["Subject"] = subject
    msg["From"] = from_addr

    body = record.get("body", "") or ""
    html = record.get("html")
    if html:
        msg.set_content(body if body else "This message contains HTML content.")
        msg.add_alternative(html, subtype="html")
    else:
        msg.set_content(body)

    attachments_field = record.get("attachments_dir") or record.get("attachments") or None
    attachments_dir = Path(str(attachments_field)) if attachments_field else default_attachments

    attach_files(msg, attachments_dir, include_pdfs=include_pdfs)
    attach_word_templates(msg, templates_dir)
    return msg


def send_all_messages(messages: List[EmailMessage]) -> None:
    server = os.environ.get("SMTP_SERVER")
    if not server:
        raise KeyError("SMTP_SERVER environment variable is required to send emails.")
    port = int(os.environ.get("SMTP_PORT", 587))
    username = os.environ.get("SMTP_USERNAME")
    password = os.environ.get("SMTP_PASSWORD")
    if not (username and password):
        raise KeyError("SMTP_USERNAME and SMTP_PASSWORD must be set to send emails.")

    with smtplib.SMTP(server, port) as smtp:
        smtp.starttls()
        smtp.login(username, password)
        for msg in messages:
            recipients = []
            for hdr in ("To", "Cc", "Bcc"):
                val = msg.get(hdr)
                if val:
                    parts = re.split(r"[,;]+", val)
                    recipients.extend([p.strip() for p in parts if p.strip()])
            logging.info("Sending to %s (subject: %s)", recipients, msg.get("Subject"))
            smtp.send_message(msg, from_addr=msg.get("From"), to_addrs=recipients)


def save_draft(msg: EmailMessage, directory: Path) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    to_safe = sanitize_filename(msg.get("To", "unknown"))
    subj_safe = sanitize_filename(msg.get("Subject", "no-subject"))
    filename = f"{to_safe}_{subj_safe}.eml"
    path = directory / filename
    i = 1
    while path.exists():
        path = directory / f"{to_safe}_{subj_safe}-{i}.eml"
        i += 1
    with path.open("wb") as fh:
        fh.write(msg.as_bytes())
    logging.info("Saved draft: %s", path)


def main(argv: Optional[Iterable[str]] = None) -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("data", type=Path, help="Path to JSON or Excel file")
    parser.add_argument(
        "--attachments",
        type=Path,
        default=Path("data/attachments"),
        help="Default directory containing files to attach",
    )
    parser.add_argument(
        "--templates",
        type=Path,
        default=None,
        help="Directory containing Word templates to attach (.docx, .doc) (optional)",
    )
    parser.add_argument(
        "--attach-pdf",
        type=str,
        choices=["Yes", "No"],
        default="Yes",
        help="Whether to attach PDFs from the attachments directory (Yes/No). Default Yes.",
    )
    parser.add_argument(
        "--sheet-name",
        type=str,
        default=None,
        help="If data is an Excel file, specify the sheet name to read (optional).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("output/drafts"),
        help="Directory to store draft emails",
    )
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--send", action="store_true", help="Send emails via SMTP")
    group.add_argument("--draft", action="store_true", help="Only save drafts (default)")
    args = parser.parse_args(argv)

    if not args.send and not args.draft:
        args.draft = True

    include_pdfs = True if args.attach_pdf == "Yes" else False

    records = load_records(args.data, sheet_name=args.sheet_name)
    if not records:
        logging.warning("No records found in %s", args.data)
        return

    messages: List[EmailMessage] = []
    for record in records:
        try:
            message = create_message(record, args.attachments, include_pdfs, args.templates)
            messages.append(message)
        except Exception as e:
            logging.error("Failed to create message for record %s: %s", record, e)

    if args.send:
        try:
            send_all_messages(messages)
        except Exception as e:
            logging.error("Failed to send messages: %s", e)
    else:
        for msg in messages:
            try:
                save_draft(msg, args.output)
            except Exception as e:
                logging.error("Failed to save draft for %s: %s", msg.get("To"), e)


if __name__ == "__main__":
    main()
