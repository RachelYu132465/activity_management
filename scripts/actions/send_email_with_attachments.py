"""Send or draft emails with attachments loaded from a directory.

This utility loads email records from a JSON or Excel file and attaches
all files in a specified directory. The resulting messages can either
be sent via SMTP or saved as draft ``.eml`` files.

Example usage::

    python scripts/actions/send_email_with_attachments.py data.json \
        --attachments data/attachments --send

Environment variables used when ``--send`` is specified:
    SMTP_SERVER, SMTP_PORT, SMTP_USERNAME, SMTP_PASSWORD
"""
from __future__ import annotations

import argparse
import importlib
import json
import os
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable, List, Dict, Any
import smtplib


def load_records(path: Path) -> List[Dict[str, Any]]:
    """Load email records from a JSON or Excel file.

    The file must contain objects with at least ``to``, ``subject`` and
    ``body`` fields. Excel support requires the ``openpyxl`` package.
    """
    ext = path.suffix.lower()
    if ext == ".json":
        with path.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
        if isinstance(data, dict):
            # support single record wrapped as dict
            data = [data]
        return data
    if ext in {".xls", ".xlsx"}:
        try:
            openpyxl = importlib.import_module("openpyxl")
        except ModuleNotFoundError as exc:
            raise ModuleNotFoundError(
                "openpyxl is required to read Excel files"
            ) from exc
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        headers = [str(h).strip().lower() for h in rows[0]]
        records: List[Dict[str, Any]] = []
        for row in rows[1:]:
            record = {headers[i]: row[i] for i in range(len(headers))}
            records.append(record)
        return records
    raise ValueError(f"Unsupported file extension: {ext}")


def attach_files(message: EmailMessage, directory: Path) -> None:
    """Attach all files from ``directory`` to ``message``."""
    for file_path in sorted(directory.glob("*")):
        if not file_path.is_file():
            continue
        with file_path.open("rb") as fh:
            data = fh.read()
        message.add_attachment(
            data,
            maintype="application",
            subtype="octet-stream",
            filename=file_path.name,
        )


def create_message(record: Dict[str, Any], attachments: Path) -> EmailMessage:
    msg = EmailMessage()
    msg["To"] = record.get("to", "")
    msg["Subject"] = record.get("subject", "")
    msg["From"] = os.environ.get("SMTP_USERNAME", "noreply@example.com")
    msg.set_content(record.get("body", ""))
    attach_files(msg, attachments)
    return msg


def send_email(msg: EmailMessage) -> None:
    server = os.environ["SMTP_SERVER"]
    port = int(os.environ.get("SMTP_PORT", 587))
    username = os.environ["SMTP_USERNAME"]
    password = os.environ["SMTP_PASSWORD"]
    with smtplib.SMTP(server, port) as smtp:
        smtp.starttls()
        smtp.login(username, password)
        smtp.send_message(msg)


def save_draft(msg: EmailMessage, directory: Path) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    filename = f"{msg['To']}_{msg['Subject']}.eml".replace("/", "-")
    with (directory / filename).open("wb") as fh:
        fh.write(bytes(msg))


def main(argv: Iterable[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("data", type=Path, help="Path to JSON or Excel file")
    parser.add_argument(
        "--attachments",
        type=Path,
        default=Path("data/attachments"),
        help="Directory containing files to attach",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("output/drafts"),
        help="Directory to store draft emails",
    )
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--send", action="store_true", help="Send emails via SMTP")
    group.add_argument(
        "--draft",
        action="store_true",
        help="Only save drafts (default)",
    )
    args = parser.parse_args(argv)

    records = load_records(args.data)
    for record in records:
        message = create_message(record, args.attachments)
        if args.send:
            send_email(message)
        else:
            save_draft(message, args.output)


if __name__ == "__main__":
    main()
