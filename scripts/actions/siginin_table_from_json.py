from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict, List


ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.actions import format_date
from scripts.actions.influencer import build_people
from scripts.actions.signin_table_render import (
    SignInDocumentContext,
    SignInRow,
    render_signin_table,
)
from scripts.core.bootstrap import DATA_DIR, OUTPUT_DIR, initialize


def load_program(program_id: int | None) -> Dict[str, Any]:
    data_file = DATA_DIR / "shared" / "program_data.json"
    programs_raw = json.loads(data_file.read_text(encoding="utf-8"))
    if isinstance(programs_raw, list):
        if program_id is not None:
            for prog in programs_raw:
                try:
                    if int(prog.get("id", -1)) == program_id:
                        return prog
                except (TypeError, ValueError):
                    continue
        return programs_raw[0] if programs_raw else {}
    elif isinstance(programs_raw, dict):
        return programs_raw
    return {}


def _get_first_nonempty(sp: dict, keys: List[str]) -> str:
    for k in keys:
        val = sp.get(k)
        if isinstance(val, str) and val.strip():
            return val.strip()
        if isinstance(val, dict):
            for subk in ("organization", "company", "affiliation", "dept", "department", "unit"):
                v2 = val.get(subk)
                if isinstance(v2, str) and v2.strip():
                    return v2.strip()
            for subk in ("title", "position", "role"):
                v2 = val.get(subk)
                if isinstance(v2, str) and v2.strip():
                    return v2.strip()
    return ""


def _safe_filename_component(value: str) -> str:
    translation = str.maketrans({ch: "_" for ch in '\\/:*?"<>|'})
    cleaned = value.translate(translation)
    return cleaned.strip() or "Program"


def main() -> None:
    parser = argparse.ArgumentParser(description="Render program speaker sign-in table")
    parser.add_argument("--program-id", type=int, default=None, help="Program id to render")
    parser.add_argument("--out", type=Path, default=None, help="Output .docx path")
    args = parser.parse_args()

    initialize()
    program = load_program(args.program_id)

    infl_file = DATA_DIR / "shared" / "influencer_data.json"
    try:
        influencers = json.loads(infl_file.read_text(encoding="utf-8"))
    except OSError:
        influencers = []
    _, speakers = build_people(program, influencers)

    program_speaker_entries = [
        entry
        for entry in (program.get("speakers") or [])
        if (entry.get("type") or "").strip() == "講者"
    ]
    program_speaker_entries.sort(key=lambda e: e.get("no", 0))

    plan_name = (program.get("planName") or "").strip()
    event_names = program.get("eventNames") or []
    event_name = (event_names[0] if event_names else "Program") or "Program"
    date = (program.get("date") or "").strip()
    slash_date = format_date(date, sep="/") if date else None

    rows: List[SignInRow] = []
    for idx, sp in enumerate(speakers):
        name_val = (sp.get("name") or "").strip()
        topic_val = ""
        if idx < len(program_speaker_entries):
            topic_val = (program_speaker_entries[idx].get("topic") or "").strip()
        title_val = _get_first_nonempty(sp, ["title", "position", "role"])
        org_val = _get_first_nonempty(
            sp,
            [
                "organization",
                "company",
                "affiliation",
                "department",
                "dept",
                "unit",
                "employer",
            ],
        )
        rows.append(
            SignInRow(
                topic=topic_val,
                name=name_val,
                title=title_val,
                organization=org_val,
            )
        )

    safe_event_name = _safe_filename_component(event_name)
    out_path = args.out or (OUTPUT_DIR / f"講師簽到表_{safe_event_name}.docx")

    context = SignInDocumentContext(
        plan_name=plan_name or event_name,
        event_name=event_name,
        date_display=slash_date,
        subtitle="講員簽到單",
    )

    result = render_signin_table(context, rows, out_path)
    print(
        f"Saved sign-in sheet to {result.output_path} "
        f"(page_width_cm={result.page_width_cm:.2f}, available_cm={result.available_width_cm:.2f}, "
        f"cols_cm={list(result.columns_cm)})"
    )


if __name__ == "__main__":
    main()
