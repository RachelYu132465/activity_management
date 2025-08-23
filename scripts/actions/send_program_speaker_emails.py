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
from typing import List, Dict, Any

# --- minimal bootstrap to allow absolute imports ---
ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# project imports
from scripts.core.build_mapping import get_event_speaker_mappings
from scripts.actions import mail_template_utils
from scripts.actions.send_email_with_attachments import (
    load_smtp_config,
    find_program_by_id,
    create_message,
    send_all_messages,
    save_draft,
)
from scripts.core.data_util import load_programs, DEFAULT_SHARED_JSON

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def build_speaker_records(program_id: str) -> List[Dict[str, Any]]:
    """Return list of speaker records merged with program info."""
    programs = load_programs(DEFAULT_SHARED_JSON)
    program = find_program_by_id(programs, program_id)

    if not program:
        raise ValueError(f"Program id {program_id} not found")

    event_names = program.get("eventNames") or []
    if not event_names:
        raise ValueError(f"Program {program_id} has no eventNames")
    event_name = event_names[0]

    mappings = get_event_speaker_mappings(event_name)
    records: List[Dict[str, Any]] = []
    for m in mappings:
        email = mail_template_utils.find_email_in_record(m)
        if not email:
            logging.warning("No email for speaker %s", m.get("name"))
            continue
        rec = dict(m)
        rec["email"] = email
        rec["program_data"] = program
        records.append(rec)
    return records


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

    load_smtp_config(Path("config/smtp.json"))

    template_path = mail_template_utils.find_template_file(args.template)
    records = build_speaker_records(args.program_id)

    messages = []
    for rec in records:
        msg = create_message(rec, template_path, [], False, None)
        messages.append(msg)
        if args.draft:
            save_draft(msg, args.output)

    if args.send and messages:
        send_all_messages(messages)

    logging.info("Prepared %d message(s)", len(messages))


if __name__ == "__main__":
    main()
