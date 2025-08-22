#!/usr/bin/env python3
"""Send email to all speakers of a program using influencer data.

Usage:
  python scripts/actions/send_program_speaker_emails.py --program-id 2 --template 我的模板.docx --send

If --send is omitted, emails are saved as .eml drafts under output/speaker_drafts.
"""
from __future__ import annotations
from pathlib import Path
import sys
import argparse
import logging
from typing import List

# --- minimal bootstrap to allow absolute imports ---
ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# project imports
from scripts.core.build_mapping import get_program_speaker_mappings
from scripts.actions import template_utils
from scripts.actions.send_email_with_attachments import (
    load_smtp_config,
    create_message,
    send_all_messages,
    save_draft,
)

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def main(argv: List[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--program-id", required=True, help="Program id to send emails for")
    parser.add_argument("--template", required=True, help="Template filename under templates/")
    parser.add_argument("--output", type=Path, default=Path("output/speaker_drafts"), help="Draft output directory")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--send", action="store_true", help="Send emails via SMTP")
    group.add_argument("--draft", action="store_true", help="Save drafts only (default)")
    args = parser.parse_args(argv)

    if not args.send and not args.draft:
        args.draft = True

    if args.send:
        try:
            load_smtp_config(Path("config/smtp.json"))
        except Exception as e:
            logging.error("Failed to load SMTP config: %s", e)
            raise

    if args.draft:
        args.output.mkdir(parents=True, exist_ok=True)

    try:
        template_path = template_utils.find_template_file(args.template)
    except Exception:
        logging.error("Template not found: %s", args.template)
        raise SystemExit(2)
    if not template_path or not template_path.exists():
        logging.error("Template not found: %s", args.template)
        raise SystemExit(2)

    try:
        records = get_program_speaker_mappings(args.program_id)
    except Exception as e:
        logging.error("Failed to load program speaker mappings for id %s: %s", args.program_id, e)
        raise SystemExit(3)

    messages = []
    for rec in records:
        email = template_utils.find_email_in_record(rec)
        if not email:
            logging.warning("No email for speaker %s", rec.get("name"))
            continue
        rec = dict(rec)
        rec["email"] = email
        msg = create_message(
            rec,
            template_path,
            attachments_entries=[],
            include_pdfs=False,
            templates_dir=None,
        )
        messages.append(msg)
        if args.draft:
            save_draft(msg, Path(args.output))

    if args.send and messages:
        send_all_messages(messages)

    logging.info("Prepared %d message(s)", len(messages))


if __name__ == "__main__":
    main()
