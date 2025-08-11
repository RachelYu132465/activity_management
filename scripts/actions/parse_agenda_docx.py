# parse_agenda_docx.py
# pip install python-docx
from __future__ import annotations

import json
import re
from pathlib import Path
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple, Iterable
from collections import Counter

from docx import Document
from docx.oxml.ns import qn

TIME_RE = re.compile(r"(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})")

TYPE_A = "TYPE_A"  # specials: merged row WITH time, or merged row WITHOUT time but title contains special keywords
TYPE_B = "TYPE_B"  # specials: merged row WITHOUT time

# Default special keywords (extendable via --special-keywords)
DEFAULT_SPECIAL_KEYWORDS = (
    "休息", "致詞", "主持人", "主持", "大合照", "合照",
    "報到", "前測", "後測", "問卷", "簽到", "簽退",
)

# Keywords indicating an END-of-day special (controls when we mark after_speaker=999)
END_SPECIAL_KEYWORDS = (
    "問卷", "後測", "後測與問卷", "問卷與後測", "結語", "Closing", "closing",
)

def _cell_gridspan(cell) -> int:
    try:
        tcPr = cell._tc.get_or_add_tcPr()  # noqa
        gridSpan = tcPr.find(qn('w:gridSpan'))
        if gridSpan is not None and gridSpan.get(qn('w:val')):
            return int(gridSpan.get(qn('w:val')))
    except Exception:
        pass
    return 1

def _clean_text(s: str) -> str:
    return (s or "").replace("\u00A0", " ").strip()

def _extract_time_range(text: str) -> Optional[Tuple[str, str]]:
    if not text:
        return None
    for line in text.splitlines():
        m = TIME_RE.search(line)
        if m:
            return (m.group(1), m.group(2))
    m = TIME_RE.search(text)
    if m:
        return (m.group(1), m.group(2))
    return None

def _parse_time_to_dt(hhmm: str) -> datetime:
    return datetime.strptime(hhmm, "%H:%M")

def _minutes_between(a: str, b: str) -> int:
    dt_a, dt_b = _parse_time_to_dt(a), _parse_time_to_dt(b)
    if dt_b < dt_a:
        dt_b += timedelta(days=1)
    return int((dt_b - dt_a).total_seconds() // 60)

def _pick_agenda_table(doc: Document):
    """Pick the first table that looks like a 3-column agenda (時間/主題/講者)."""
    def norm(s: str) -> str:
        return _clean_text(s).replace(" ", "").replace("\n", "")
    for tbl in doc.tables:
        if len(tbl.columns) >= 3 and len(tbl.rows) >= 2:
            headers = []
            for i in range(min(3, len(tbl.rows))):
                cells = tbl.rows[i].cells
                if len(cells) >= 3:
                    headers.append((norm(cells[0].text), norm(cells[1].text), norm(cells[2].text)))
            header_text = " ".join(["|".join(h) for h in headers])
            if ("時間" in header_text) and ("主題" in header_text or "主旨" in header_text) and ("講者" in header_text or "講師" in header_text):
                return tbl
    return doc.tables[0] if doc.tables else None

def _looks_like_special_text(topic_text: str, speaker_text: str, keywords: Iterable[str]) -> bool:
    hay = f"{_clean_text(topic_text)}\n{_clean_text(speaker_text)}"
    return any(k in hay for k in keywords)

def parse_agenda(
        docx_path: Path,
        event_name: str,
        template: str = TYPE_A,
        special_keywords: Iterable[str] = DEFAULT_SPECIAL_KEYWORDS,
) -> Dict[str, Any]:
    doc = Document(docx_path)
    tbl = _pick_agenda_table(doc)
    if tbl is None:
        raise SystemExit("找不到任何表格。請確認檔案內含有 3 欄的議程表。")

    # Skip header row if matches (時間/主題/講者)
    start_row_idx = 0
    if len(tbl.rows) >= 1:
        heads = [_clean_text(c.text) for c in tbl.rows[0].cells[:3]]
        if any("時間" in h for h in heads) and any(("主題" in h or "主旨" in h) for h in heads) and any(("講者" in h or "講師" in h) for h in heads):
            start_row_idx = 1

    talks: List[Dict[str, Any]] = []
    specials: List[Dict[str, Any]] = []
    special_times: List[Optional[Tuple[str, str]]] = []  # align with specials

    # Event timeline to know what's the very last row and its type
    timeline: List[Tuple[str, int]] = []  # ("talk", talk_no) or ("special", special_index)

    first_time_start: Optional[str] = None
    last_time_end: Optional[str] = None
    prev_is_talk_idx = 0
    talk_durations: List[int] = []

    for r in tbl.rows[start_row_idx:]:
        cells = r.cells
        if len(cells) < 3:
            # Fully merged single cell row => TYPE_B special candidate
            row_text = _clean_text("\n".join([c.text for c in cells]))
            span_total = sum(_cell_gridspan(c) for c in cells)
            merged_across = (len(cells) == 1 and _cell_gridspan(cells[0]) >= 3) or (span_total >= 3 and len(cells) == 1)
            if template == TYPE_B and merged_across:
                sidx = len(specials)
                specials.append({
                    "after_speaker": prev_is_talk_idx if prev_is_talk_idx > 0 else 0,
                    "title": row_text,
                    "duration": None
                })
                special_times.append(None)
                timeline.append(("special", sidx))
            continue

        time_text = _clean_text(cells[0].text)
        topic_text = _clean_text(cells[1].text)
        speaker_text = _clean_text(cells[2].text)

        # Detect merged across topic/speaker
        topic_span = _cell_gridspan(cells[1])
        speaker_span = _cell_gridspan(cells[2])
        row_span_total = _cell_gridspan(cells[0]) + topic_span + speaker_span
        is_merged_across_ts = (topic_span >= 2) or (row_span_total >= 3 and (topic_span >= 2 or speaker_span >= 2))

        time_rng = _extract_time_range(time_text)

        # TYPE_A: merged AND (has time OR contains any special keywords like 休息/致詞/主持人/大合照…)
        if template == TYPE_A:
            is_special = is_merged_across_ts and (time_rng is not None or _looks_like_special_text(topic_text, speaker_text, special_keywords))
        elif template == TYPE_B:
            is_special = is_merged_across_ts and (time_rng is None)
        else:
            raise SystemExit(f"未知的 template 類型：{template}（請用 {TYPE_A}/{TYPE_B}）")

        if is_special:
            duration_val: Optional[int] = None
            if time_rng is not None:
                duration_val = _minutes_between(time_rng[0], time_rng[1])
            sidx = len(specials)
            specials.append({
                "after_speaker": prev_is_talk_idx if prev_is_talk_idx > 0 else 0,
                "title": topic_text if topic_text else (speaker_text or ""),
                "duration": duration_val
            })
            special_times.append(time_rng)
            timeline.append(("special", sidx))
            # Only update day start/end if this special has time
            if time_rng is not None:
                if first_time_start is None:
                    first_time_start = time_rng[0]
                last_time_end = time_rng[1]
            continue

        # Talk row: needs a time; if time is in topic cell, try that
        if time_rng is None:
            time_rng = _extract_time_range(topic_text)
        if time_rng is None:
            continue

        if first_time_start is None:
            first_time_start = time_rng[0]
        last_time_end = time_rng[1]

        prev_is_talk_idx += 1
        talk_minutes = _minutes_between(time_rng[0], time_rng[1])
        talk_durations.append(talk_minutes)
        talks.append({
            "no": prev_is_talk_idx,
            "topic": topic_text if topic_text else "(未命名主題)",
            "name": speaker_text,
        })
        timeline.append(("talk", prev_is_talk_idx))

    # Only mark the LAST ROW as after_speaker=999 if it is a special and it looks like an end-of-day item
    if timeline and timeline[-1][0] == "special":
        sidx = timeline[-1][1]
        title = specials[sidx]["title"]
        tr = special_times[sidx]
        end_matches_day_end = (tr is not None and last_time_end is not None and tr[1] == last_time_end)
        if end_matches_day_end or any(k in title for k in END_SPECIAL_KEYWORDS):
            specials[sidx]["after_speaker"] = 999

    # Use the mode of talk durations for speaker_minutes
    speaker_minutes = 50
    if talk_durations:
        speaker_minutes = Counter(talk_durations).most_common(1)[0][0]

    if first_time_start is None:
        first_time_start = "09:00"
    if last_time_end is None:
        last_time_end = "17:00"

    data = {
        "eventNames": [event_name],
        "agenda_settings": {
            "start_time": first_time_start,
            "end_time": last_time_end,
            "speaker_minutes": speaker_minutes,
            "special_sessions": specials
        },
        "speakers": talks,
        "activities_contacts": [{
            "ID": "",
            "id_number": "",
            "name": "",
            "gender": "",
            "title": "",
            "unity": [],
            "email": "",
            "contact_person": {"name": "", "email": ""},
            "group_by_program": [],
            "name_mail_use": ""
        }]
    }
    return data

if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Parse a 3-column agenda DOCX (時間/主題/講者) into activities_data JSON.")
    ap.add_argument("--docx", required=True, help="Path to the agenda .docx")
    ap.add_argument("--event-name", required=True, help="Event name to put into eventNames")
    ap.add_argument("--template", default=TYPE_A, choices=[TYPE_A, TYPE_B], help="Agenda template TYPE_A or TYPE_B")
    ap.add_argument("--out", required=True, help="Output JSON file path")
    ap.add_argument("--special-keywords", default=",".join(DEFAULT_SPECIAL_KEYWORDS),
                    help="以逗號分隔的 special 關鍵字清單（預設：休息,致詞,主持人,主持,大合照,合照,報到,前測,後測,問卷,簽到,簽退）")
    args = ap.parse_args()

    keywords = tuple([k for k in (s.strip() for s in args.special_keywords.split(",")) if k])
    payload = [parse_agenda(Path(args.docx), args.event_name, template=args.template, special_keywords=keywords)]
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[OK] Wrote {out_path}")
