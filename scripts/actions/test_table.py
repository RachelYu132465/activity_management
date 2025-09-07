from __future__ import annotations
from pathlib import Path
import json
import sys


ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.core.bootstrap import DATA_DIR

from typing import Any, Dict, List

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

program = load_program(5)
json.loads((DATA_DIR / "shared/program_data.json").read_text(encoding="utf-8"))
plan_name = program.get("planName")
print(plan_name)

# # test_table.py
# from __future__ import annotations
# from docx import Document
# from docx.enum.table import WD_TABLE_ALIGNMENT
# from docx.enum.table import WD_ROW_HEIGHT_RULE
# from docx.enum.text import WD_ALIGN_PARAGRAPH
# from docx.oxml import OxmlElement
# from docx.oxml.ns import qn
# from docx.shared import Cm, Pt
# import sys
#
# # ----- 參數（直接在程式碼改） -----
# OUT_PATH = "test_table.docx"
# LEFT_RIGHT_MARGIN_CM = 1.5   # 左右邊界 (cm)
# COL_WIDTHS_CM = [10.0, 6.0, 6.0]  # 三欄寬（cm），可改（但總和若超過可用寬度會自動縮放）
# DATA_ROWS = 8               # 要建立的空白資料列數（不包含表頭）
# HDR_HEIGHT_CM = 0.9         # 表頭列高 (cm)
# DATA_ROW_HEIGHT_CM = 1.0    # 每筆資料列高 (cm)
# CELL_PADDING_CM = 0.12      # cell 左右內距 (cm)
# FONT_PT = 11                # 測試時用的字型大小 (pt)
# # ----------------------------------
#
# def set_table_cell_margins(table, left_cm: float = 0.1, right_cm: float = 0.1, top_pt: int = 0, bottom_pt: int = 0):
#     tbl = table._tbl
#     tblPr = tbl.tblPr
#     tcMar = tblPr.find(qn("w:tblCellMar"))
#     if tcMar is None:
#         tcMar = OxmlElement("w:tblCellMar")
#         tblPr.append(tcMar)
#     def _set_node(name: str, cm_val: float, parent: OxmlElement):
#         node = parent.find(qn(f"w:{name}"))
#         if node is None:
#             node = OxmlElement(f"w:{name}")
#             parent.append(node)
#         node.set(qn("w:w"), str(int(cm_val * 567)))  # dxa
#         node.set(qn("w:type"), "dxa")
#     _set_node("left", left_cm, tcMar)
#     _set_node("right", right_cm, tcMar)
#     # top/bottom as dxa, keep 0 if not provided
#     _set_node("top", (top_pt / 72.0) * 2.54 if top_pt else 0, tcMar)
#     _set_node("bottom", (bottom_pt / 72.0) * 2.54 if bottom_pt else 0, tcMar)
#
# def _set_table_total_width(table, total_cm: float):
#     tbl = table._tbl
#     tblPr = tbl.tblPr
#     existing = tblPr.find(qn("w:tblW"))
#     if existing is not None:
#         tblPr.remove(existing)
#     tblW = OxmlElement("w:tblW")
#     tblW.set(qn("w:w"), str(int(total_cm * 567)))
#     tblW.set(qn("w:type"), "dxa")
#     tblPr.append(tblW)
#
# def safe_set_row_height(row, height_cm: float, preferred_rule: str = "EXACT"):
#     # 設定高度
#     row.height = Cm(height_cm)
#     # 安全地設定 height_rule（不同版本 python-docx 的 enum 可能不一樣）
#     rule_value = None
#     if hasattr(WD_ROW_HEIGHT_RULE, preferred_rule):
#         rule_value = getattr(WD_ROW_HEIGHT_RULE, preferred_rule)
#     else:
#         for alt in ("EXACT", "AT_LEAST", "AUTO"):
#             if hasattr(WD_ROW_HEIGHT_RULE, alt):
#                 rule_value = getattr(WD_ROW_HEIGHT_RULE, alt)
#                 break
#     if rule_value is not None:
#         try:
#             row.height_rule = rule_value
#         except Exception:
#             pass
#
# def main():
#     doc = Document()
#
#     # 設左右邊界（確保 A4 可用區域）
#     doc.sections[0].left_margin = Cm(LEFT_RIGHT_MARGIN_CM)
#     doc.sections[0].right_margin = Cm(LEFT_RIGHT_MARGIN_CM)
#
#     # 建 table
#     tbl = doc.add_table(rows=1, cols=3, style="Table Grid")
#     tbl.autofit = False
#     tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
#     try:
#         tbl.left_indent = Cm(0)
#     except Exception:
#         pass
#
#     # 計算可用寬度與欄寬調整保護
#     page_width_cm = float(doc.sections[0].page_width) / float(Cm(1))
#     available_cm = page_width_cm - (LEFT_RIGHT_MARGIN_CM * 2)
#
#     # 若 COL_WIDTHS_CM 總和 > available_cm，等比例縮放
#     total_requested = sum(COL_WIDTHS_CM)
#     col_widths = COL_WIDTHS_CM.copy()
#     if total_requested > available_cm:
#         scale = available_cm / total_requested
#         col_widths = [w * scale for w in col_widths]
#
#     # assign widths
#     for i, w in enumerate(col_widths):
#         tbl.columns[i].width = Cm(w)
#
#     # 強制 table 總寬寫入 xml（避免 Word 自動伸張）
#     _set_table_total_width(tbl, sum(col_widths))
#
#     # 縮小 cell padding
#     set_table_cell_margins(tbl, left_cm=CELL_PADDING_CM, right_cm=CELL_PADDING_CM)
#
#     # 表頭（放一個可見文字）
#     hdr = tbl.rows[0]
#     safe_set_row_height(hdr, HDR_HEIGHT_CM)
#     hdr.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
#     hdr.cells[0].paragraphs[0].add_run("Topic / 主題").font.size = Pt(FONT_PT)
#     hdr.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
#     hdr.cells[1].paragraphs[0].add_run("Speaker / 講者").font.size = Pt(FONT_PT)
#     hdr.cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
#     hdr.cells[2].paragraphs[0].add_run("Sign-in / 簽到").font.size = Pt(FONT_PT)
#
#     # 建幾列空白測試列（每個 cell 放 NBSP 保持不空）
#     for _ in range(DATA_ROWS):
#         r = tbl.add_row()
#         safe_set_row_height(r, DATA_ROW_HEIGHT_CM)
#         for c in r.cells:
#             p = c.paragraphs[0]
#             p.alignment = WD_ALIGN_PARAGRAPH.LEFT
#             # non-breaking space 保持 cell 可見，便於檢查 size
#             run = p.add_run("\u00A0")
#             run.font.size = Pt(FONT_PT)
#
#     doc.save(OUT_PATH)
#     print(f"Saved test file to {OUT_PATH} - page width: {page_width_cm:.2f} cm, available: {available_cm:.2f} cm, cols: {col_widths}")
#
# if __name__ == "__main__":
#     main()
