#!/usr/bin/env python3
"""Render program handbook to a Word document.

This script loads ``program_data.json`` and ``influencer_data.json`` from the
``data/shared`` directory, selects a program by ``--program-id`` (defaults to the
first program) and creates a simple `.docx` file in the ``output`` directory.

The goal is to mirror ``templates/template.html`` but for Word output.  The
layout is intentionally simple so that the generated document remains
readable even without HTML rendering support.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict, List

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

# Project helpers
from scripts.core.bootstrap import DATA_DIR, OUTPUT_DIR, initialize
from scripts.actions.influencer import build_people


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def load_program(program_id: int | None) -> Dict[str, Any]:
    """Return the program matching ``program_id`` from program_data.json.

    If ``program_id`` is ``None`` or not found, the first program entry is
    returned.
    """
    data_file = DATA_DIR / "shared" / "program_data.json"
    programs_raw = json.loads(data_file.read_text(encoding="utf-8"))

    # program_data.json may contain a list or a single dict
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


def build_schedule(event: Dict[str, Any]) -> List[Dict[str, str]]:
    """Build schedule rows from speaker information.

    The result is a list of dictionaries containing ``time``, ``topic`` and
    ``speaker`` keys, using each speaker's ``start_time`` and ``end_time``
    fields.  Special sessions are ignored so that only the regular speaker
    timetable is returned.
    """
    def time_range(start: str | None, end: str | None) -> str:
        if start and end:
            return f"{start}-{end}"
        return start or end or ""

    speakers = event.get("speakers", []) or []

    rows: List[Dict[str, str]] = []

    host = next((sp for sp in speakers if sp.get("type") == "主持人"), None)
    if host:
        text = " ".join(
            filter(
                None,
                [
                    time_range(host.get("start_time"), host.get("end_time")),
                    host.get("topic"),
                    host.get("name"),
                ],
            )
        )
        rows.append({"kind": "host", "time": "", "topic": text, "speaker": ""})

    for sp in speakers:
        if sp.get("type") == "主持人":
            continue
        start = sp.get("start_time")
        end = sp.get("end_time")
        rows.append(
            {
                "kind": "talk",
                "time": time_range(start, end),
                "topic": sp.get("topic", ""),
                "speaker": sp.get("name", ""),
            }
        )
    return rows


def set_run_font(run, size_pt: int, bold: bool = False) -> None:
    """Apply project font settings to ``run``.

    Chinese characters use Microsoft JhengHei while Latin characters use
    Times New Roman.  ``size_pt`` is the font size in points.
    """
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft JhengHei")
    run.font.size = Pt(size_pt)
    run.bold = bold


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Render program to docx")
    parser.add_argument("--program-id", type=int, default=None, help="Program id to render")
    parser.add_argument("--out", type=Path, default=None, help="Output .docx path")
    args = parser.parse_args()

    initialize()
    program = load_program(args.program_id)

    # Build chairs/speakers enriched with influencer data
    infl_file = DATA_DIR / "shared" / "influencer_data.json"
    try:
        influencers = json.loads(infl_file.read_text(encoding="utf-8"))
    except OSError:
        influencers = []
    chairs, speakers = build_people(program, influencers)

    schedule_rows = build_schedule(program)

    event_name = (program.get("eventNames") or ["Program"])[0]

    out_path = args.out or (OUTPUT_DIR / f"program_{program.get('id', '0')}.docx")

    doc = Document()
    normal_style = doc.styles["Normal"]
    normal_font = normal_style.font
    normal_font.name = "Times New Roman"
    normal_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft JhengHei")
    normal_font.size = Pt(12)

    # Heading style for TOC inclusion
    heading1 = doc.styles["Heading 1"]
    heading1.font.name = "Times New Roman"
    heading1._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft JhengHei")
    heading1.font.size = Pt(16)
    heading1.font.bold = True

    # Style constants (pt) derived from the HTML template
    TITLE_PT = 28
    SECTION_PT = 16
    NAME_PT = 18
    PROFILE_PT = 14
    TABLE_PT = 11

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run(event_name)
    set_run_font(title_run, TITLE_PT, bold=True)


    # Cover details
    cover_lines = []
    if program.get("date"):
        cover_lines.append(f"日期：{program['date']}")
    locations = program.get("locations") or []
    if locations:
        loc_text = locations[0]
        if len(locations) > 1:
            loc_text += f"（{locations[1]}）"
        cover_lines.append(f"地點：{loc_text}")
    if program.get("organizers"):
        cover_lines.append(
            f"主辦單位：{'、'.join(program['organizers'])}"
        )
    if program.get("coOrganizers"):
        cover_lines.append(
            f"協辦單位：{'、'.join(program['coOrganizers'])}"
        )
    if program.get("jointOrganizers"):
        cover_lines.append(
            f"合辦單位：{'、'.join(program['jointOrganizers'])}"
        )
    for line in cover_lines:
        p = doc.add_paragraph()
        run = p.add_run(line)
        set_run_font(run, PROFILE_PT, bold=True)

    doc.add_page_break()

    # Table of contents
    toc_title_p = doc.add_paragraph()
    toc_title_run = toc_title_p.add_run("目錄")
    set_run_font(toc_title_run, SECTION_PT, bold=True)

    doc.add_page_break()

    # Table of contents
    doc.add_heading("目錄", level=1)

    toc_p = doc.add_paragraph()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), 'TOC \\o "1-3" \\h \\z \\u')
    toc_p._p.append(fld)
    doc.add_page_break()

    # Activity info section
    doc.add_heading("活動資訊", level=1)
    info_lines = []
    if program.get("date"):
        info_lines.append(f"日期：{program['date']}")
    if locations:
        info_lines.append(f"地點：{loc_text}")
    if program.get("organizers"):
        info_lines.append(f"主辦單位：{'、'.join(program['organizers'])}")
    if program.get("coOrganizers"):
        info_lines.append(f"協辦單位：{'、'.join(program['coOrganizers'])}")
    if program.get("jointOrganizers"):
        info_lines.append(f"合辦單位：{'、'.join(program['jointOrganizers'])}")
    for line in info_lines:
        p = doc.add_paragraph(line)
        for r in p.runs:
            set_run_font(r, PROFILE_PT)
    doc.add_page_break()


    doc.add_heading("議程", level=1)
    if schedule_rows:
        table = doc.add_table(rows=1, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr = table.rows[0].cells
        headers = ["時間", "議程", "講者"]
        for idx, text in enumerate(headers):
            p = hdr[idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(text)
            set_run_font(run, TABLE_PT, bold=True)
        for row in schedule_rows:
            cells = table.add_row().cells
            data = [row.get("time", ""), row.get("topic", ""), row.get("speaker", "")]
            for idx, text in enumerate(data):
                p = cells[idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(text)
                set_run_font(run, TABLE_PT)
    doc.add_page_break()

    if chairs:
        doc.add_heading("主持人", level=1)
        for ch in chairs:
            p = doc.add_paragraph()
            name_run = p.add_run(ch.get("name", ""))
            set_run_font(name_run, NAME_PT, bold=True)
            title = ch.get("title")
            if title:
                title_run = p.add_run(f" {title}")
                set_run_font(title_run, NAME_PT)
            prof = ch.get("profile")
            if prof:
                prof_p = doc.add_paragraph(prof)
                for r in prof_p.runs:
                    set_run_font(r, PROFILE_PT)
        doc.add_page_break()

    if speakers:
        doc.add_heading("講者", level=1)
        for sp in speakers:
            p = doc.add_paragraph()
            name_run = p.add_run(sp.get("name", ""))
            set_run_font(name_run, NAME_PT, bold=True)
            title = sp.get("title")
            if title:
                title_run = p.add_run(f" {title}")
                set_run_font(title_run, NAME_PT)
            prof = sp.get("profile")
            if prof:
                prof_p = doc.add_paragraph(prof)
                for r in prof_p.runs:
                    set_run_font(r, PROFILE_PT)

    # Footer page numbers (skip cover page)
    section = doc.sections[0]
    section.different_first_page_header_footer = True
    footer_p = section.footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_p.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_begin, instr, fld_end])
    set_run_font(run, 12)

    doc.save(out_path)
    print(f"Saved docx to {out_path}")


if __name__ == "__main__":
    main()
