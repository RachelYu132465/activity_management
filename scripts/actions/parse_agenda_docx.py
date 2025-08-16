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

TIME_RE = re.compile(r"(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})")

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

# ---- Special 正規化規則（同義字 → 統一標題）----
# 順序很重要，先比對到的就用那個標題
SPECIAL_RULES: List[Tuple[re.Pattern, str, bool]] = [
    # (pattern, normalized_title, is_end_of_day)
    (re.compile(r"(問卷|後測|後測與問卷|問卷與後測|closing|結語)", re.I), "問卷與後測", True),
    (re.compile(r"(午餐|午休|午間)", re.I),                           "午間休息", False),
    (re.compile(r"(大合照|合照)", re.I),                               "大合照",   False),
    (re.compile(r"(致詞|開場)", re.I),                                 "致詞",     False),
    (re.compile(r"(休息|中場休息)", re.I),                             "課間休息", False),
    (re.compile(r"(報到|簽到|簽退)", re.I),                            "報到／簽到簽退", False),
    # 需要可繼續加在這裡
]
END_TITLES = { title for _, title, is_end in SPECIAL_RULES if is_end }

def _match_special_title(topic_text: str, speaker_text: str) -> Tuple[Optional[str], bool]:
    """回傳 (normalized_title, is_end_of_day)；若無匹配則 (None, False)"""
    hay = f"{_clean_text(topic_text)}\n{_clean_text(speaker_text)}"
    for pat, norm_title, is_end in SPECIAL_RULES:
        if pat.search(hay):
            return norm_title, is_end
    return None, False

def parse_agenda(docx_path: Path, event_name: str) -> Dict[str, Any]:
    doc = Document(docx_path)
    tbl = _pick_agenda_table(doc)
    if tbl is None:
        raise SystemExit("找不到任何表格。請確認檔案內含有 3 欄的議程表。")

    # 如果第一列是表頭（時間/主題/講者），就跳過
    start_row_idx = 0
    if len(tbl.rows) >= 1:
        heads = [_clean_text(c.text) for c in tbl.rows[0].cells[:3]]
        if any("時間" in h for h in heads) and any(("主題" in h or "主旨" in h) for h in heads) and any(("講者" in h or "講師" in h) for h in heads):
            start_row_idx = 1

    talks: List[Dict[str, Any]] = []
    specials: List[Dict[str, Any]] = []
    special_times: List[Optional[Tuple[str, str]]] = []  # 與 specials 對齊
    timeline: List[Tuple[str, int]] = []  # ("talk", talk_no) or ("special", idx)

    first_time_start: Optional[str] = None
    last_time_end: Optional[str] = None
    talk_idx = 0
    talk_durations: List[int] = []

    for r in tbl.rows[start_row_idx:]:
        cells = r.cells
        if len(cells) < 3:
            # 不規則列，忽略
            continue

        time_text = _clean_text(cells[0].text)
        topic_text = _clean_text(cells[1].text)
        speaker_text = _clean_text(cells[2].text)

        time_rng = _extract_time_range(time_text)
        norm_title, is_end_flag = _match_special_title(topic_text, speaker_text)

        if norm_title is not None:
            # special（不看合併欄位、有沒有時間都算）
            duration_val: Optional[int] = None
            if time_rng is not None:
                duration_val = _minutes_between(time_rng[0], time_rng[1])
                if first_time_start is None:
                    first_time_start = time_rng[0]
                last_time_end = time_rng[1]

            sidx = len(specials)
            specials.append({
                "after_speaker": talk_idx if talk_idx > 0 else 0,
                "title": norm_title,
                "duration": duration_val
            })
            special_times.append(time_rng)
            timeline.append(("special", sidx))
            continue

        # 否則視為 talk（需要時間；若時間寫在主題欄也試著抓）
        if time_rng is None:
            time_rng = _extract_time_range(topic_text)
        if time_rng is None:
            # 沒時間：避免把主持人/備註當 talk
            continue

        if first_time_start is None:
            first_time_start = time_rng[0]
        last_time_end = time_rng[1]

        talk_idx += 1
        tmin = _minutes_between(time_rng[0], time_rng[1])
        talk_durations.append(tmin)
        talks.append({
            "no": talk_idx,
            "topic": topic_text if topic_text else "(未命名主題)",
            "name": speaker_text,
        })
        timeline.append(("talk", talk_idx))

    # 只在「最後一列是 special 且像真正收尾」才把 after_speaker 設 999
    if timeline and timeline[-1][0] == "special":
        sidx = timeline[-1][1]
        title = specials[sidx]["title"]
        tr = special_times[sidx]
        end_matches_day_end = (tr is not None and last_time_end is not None and tr[1] == last_time_end)
        if end_matches_day_end or (title in END_TITLES):
            specials[sidx]["after_speaker"] = 999

    # 講者時長取眾數
    speaker_minutes = 50
    if talk_durations:
        speaker_minutes = Counter(talk_durations).most_common(1)[0][0]

    if first_time_start is None:
        first_time_start = "09:00"
    if last_time_end is None:
        last_time_end = "17:00"

    agenda_settings = {
        "start_time": first_time_start,
        "end_time": last_time_end,
        "speaker_minutes": speaker_minutes,
        "special_sessions": specials,
    }
    activity = {
        "eventNames": [event_name],
        # keep speaker details with activities; do not move to program data
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
            "name_mail_use": "",
        }],
    }
    program_agenda = {
        "eventNames": [event_name],
        "agenda_settings": agenda_settings,
    }
    return activity, program_agenda

if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Parse a 3-column agenda DOCX (時間/主題/講者) into activities and program JSON.")
    ap.add_argument("--docx", required=True, help="Path to the agenda .docx")
    ap.add_argument("--event-name", required=True, help="Event name to put into eventNames")
    ap.add_argument("--out", required=True, help="Output activities JSON file path")
    ap.add_argument("--program-out", help="Output program JSON file path")
    args = ap.parse_args()

    activity, program = parse_agenda(Path(args.docx), args.event_name)
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps([activity], ensure_ascii=False, indent=2), encoding="utf-8")
    if args.program_out:
        prog_path = Path(args.program_out)
        prog_path.parent.mkdir(parents=True, exist_ok=True)
        prog_path.write_text(json.dumps(program, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[OK] Wrote {out_path}")
