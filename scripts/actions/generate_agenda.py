# scripts/actions/generate_agenda_docx.py
# pip install python-docx
from __future__ import annotations
import json
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Tuple

from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn

from scripts.core.bootstrap import (
    initialize, DATA_DIR, OUTPUT_DIR
)

initialize()

# ---------- data loading ----------
def load_activities() -> List[Dict[str, Any]]:
    p = DATA_DIR / "activities" / "activities_data.json"
    return json.loads(p.read_text(encoding="utf-8"))

def pick_event(activities: List[Dict[str, Any]], event_name: str) -> Dict[str, Any]:
    for ev in activities:
        names = ev.get("eventNames") or []
        if any(event_name == n for n in names):
            return ev
    raise SystemExit(f"找不到 event：{event_name}")

# ---------- time helpers ----------
def distribute_empty_durations(cfg: Dict[str, Any], speaker_count: int) -> Dict[str, Any]:
    fmt = "%H:%M"
    total_minutes = int(
        (datetime.strptime(cfg["end_time"], fmt) - datetime.strptime(cfg["start_time"], fmt)).total_seconds() / 60
    )
    total_known = speaker_count * int(cfg["speaker_minutes"])
    empty_items = []

    for s in cfg.get("special_sessions", []):
        d = s.get("duration")
        if d is None:
            empty_items.append(s)
        else:
            total_known += int(d)

    remaining = total_minutes - total_known
    if remaining < 0:
        print(f"[警告] 總時數不足，缺 {-remaining} 分鐘（請調整 speaker_minutes 或 special_sessions）")
        remaining = 0

    if empty_items:
        avg = round(remaining / len(empty_items)) if len(empty_items) else 0
        for s in empty_items:
            s["duration"] = max(0, avg)
        print(f"[INFO] 自動分配空白時段：總剩餘 {remaining} 分鐘，平均每段 {avg} 分鐘")
    return cfg

def gen_agenda_rows(event: Dict[str, Any]) -> List[Dict[str, str]]:
    cfg = distribute_empty_durations(dict(event["agenda_settings"]), len(event["speakers"]))
    start_str = cfg["start_time"]
    current = datetime.strptime(start_str, "%H:%M")

    def add_minutes(dt, m): return dt + timedelta(minutes=int(m))

    rows: List[Dict[str, str]] = []

    # 開場 special（after_speaker = 0）
    for s in cfg["special_sessions"]:
        if int(s["after_speaker"]) == 0:
            end = add_minutes(current, s["duration"])
            rows.append({
                "kind": "special",
                "time": f"{current.strftime('%H:%M')}-{end.strftime('%H:%M')}",
                "title": s["title"],
                "speaker": ""
            })
            current = end

    # 每一位講者 + 之後可能的 special
    for sp in event["speakers"]:
        end = add_minutes(current, cfg["speaker_minutes"])
        rows.append({
            "kind": "talk",
            "time": f"{current.strftime('%H:%M')}-{end.strftime('%H:%M')}",
            "title": sp["topic"],
            "speaker": sp.get("name", "")
        })
        current = end

        for s in cfg["special_sessions"]:
            if int(s["after_speaker"]) == int(sp["no"]):
                end2 = add_minutes(current, s["duration"])
                rows.append({
                    "kind": "special",
                    "time": f"{current.strftime('%H:%M')}-{end2.strftime('%H:%M')}",
                    "title": s["title"],
                    "speaker": ""
                })
                current = end2

    # 尾端 special（after_speaker = 999）
    for s in cfg["special_sessions"]:
        if int(s["after_speaker"]) == 999:
            end = add_minutes(current, s["duration"])
            rows.append({
                "kind": "special",
                "time": f"{current.strftime('%H:%M')}-{end.strftime('%H:%M')}",
                "title": s["title"],
                "speaker": ""
            })
            current = end

    return rows

# ---------- docx helpers ----------
# 顏色常數：tuple 格式 (R, G, B)
GREEN: Tuple[int, int, int] = (0x12, 0x6E, 0x2E)   # 深綠
WHITE: Tuple[int, int, int] = (0xFF, 0xFF, 0xFF)

def ensure_page_setup(doc: Document):
    """A4、窄邊界，盡量保證單頁"""
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width  = Cm(21.0)
    section.top_margin    = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin   = Cm(1.5)
    section.right_margin  = Cm(1.5)

def set_cell_shading(cell, fill_rgb_tuple: Tuple[int, int, int]):
    """設定表格儲存格底色（用 6 碼 hex）"""
    r, g, b = fill_rgb_tuple
    hexcolor = f"{r:02X}{g:02X}{b:02X}"

   table.style = "Table Grid"
def set_cell_text(cell, text: str, bold: bool = False,
                  color: Tuple[int, int, int] | None = None,
                  align=WD_ALIGN_PARAGRAPH.LEFT, size_pt: float = 10.5):
    """設定表格儲存格文字屬性"""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(text)
    run.font.size = Pt(size_pt)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)  # tuple → RGBColor
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)

def add_agenda_table(doc: Document, rows: List[Dict[str, str]], title: str | None = None):
    # 標題（可選）
    if title:
        h = doc.add_paragraph(title)
        h.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = h.runs[0] if h.runs else h.add_run()
        run.font.size = Pt(14)
        run.bold = True

    # 建表（兩欄：時間、內容）
    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"   # ← 這行
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # 欄寬：A4 可用寬約 18cm（左右各 1.5cm 邊界）
    time_w = Cm(3.6)
    body_w = Cm(18.0 - 3.6)  # ≈14.4cm
    table.columns[0].width = time_w
    table.columns[1].width = body_w

    for r in rows:
        tr = table.add_row().cells
        # 時間欄
        set_cell_text(tr[0], r["time"], bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=10.5)

        if r["kind"] == "special":
            # 綠底白字，置中
            set_cell_shading(tr[0], GREEN)
            set_cell_shading(tr[1], GREEN)
            set_cell_text(tr[0], r["time"], bold=True, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_text(tr[1], r["title"], bold=True, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
        else:
            # 內容：主題 + 換行 + 講者
            topic = r["title"].strip()
            name  = r["speaker"].strip()
            content = topic if not name else f"{topic}\n{name}"
            set_cell_text(tr[1], content, align=WD_ALIGN_PARAGRAPH.LEFT)

    # 邊框（細線）
    tbl_pr = table._tbl.get_or_add_tblPr()
    tbl_borders = OxmlElement('w:tblBorders')
    for tag in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{tag}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '6')       # 0.5pt
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), '000000')
        tbl_borders.append(el)
    tbl_pr.append(tbl_borders)

def export_agenda_docx(event: Dict[str, Any], out_path: Path):
    rows = gen_agenda_rows(event)
    doc = Document()
    ensure_page_setup(doc)

    # 標題可用第一個 eventName
    title = (event.get("eventNames") or ["議程"])[0]
    add_agenda_table(doc, rows, title=title)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    print(f"[OK] 已輸出：{out_path}")

# ---------- main ----------
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="依 eventName 讀取 activities_data.json，輸出一頁 A4 的議程表 DOCX")
    ap.add_argument("--event", required=True, help="event name（須與 activities_data.json 的 eventNames 完全相同）")
    ap.add_argument("--out", default=str(OUTPUT_DIR / "letters" / "agenda.docx"), help="輸出路徑")
    args = ap.parse_args()

    activities = load_activities()
    event = pick_event(activities, args.event)
    export_agenda_docx(event, Path(args.out))
