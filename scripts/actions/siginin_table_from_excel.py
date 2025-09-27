from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd


ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.actions import format_date
from scripts.actions.signin_table_render import (
    SignInDocumentContext,
    SignInRow,
    render_signin_table,
)
from scripts.core.bootstrap import OUTPUT_DIR, initialize


def _safe_filename_component(value: str) -> str:
    translation = str.maketrans({ch: "_" for ch in '\\/:*?"<>|'})
    cleaned = value.translate(translation)
    return cleaned.strip() or "Program"


def _detect_column(columns: Iterable[str], keywords: Iterable[str]) -> Optional[str]:
    keyword_list = [k.strip().lower() for k in keywords]
    for col in columns:
        label = str(col or "").strip()
        lower = label.lower()
        for key in keyword_list:
            if key and key in lower:
                return label
            if key and key in label:
                return label
    return None


def _strip_value(val) -> str:
    if val is None:
        return ""
    text = str(val)
    return text.strip()


def main() -> None:
    parser = argparse.ArgumentParser(description="Render speaker sign-in table from Excel")
    parser.add_argument("excel", type=Path, help="Input Excel file path")
    parser.add_argument("--sheet", help="Sheet name or index", default=0)
    parser.add_argument("--plan-name", help="Plan name for the document", default=None)
    parser.add_argument("--event-name", help="Event name for the document", default=None)
    parser.add_argument("--date", help="Event date (YYYY-MM-DD)", default=None)
    parser.add_argument("--out", type=Path, help="Output .docx path", default=None)
    parser.add_argument("--topic-col", help="Explicit topic column name", default=None)
    parser.add_argument("--name-col", help="Explicit name column name", default=None)
    parser.add_argument("--title-col", help="Explicit title column name", default=None)
    parser.add_argument("--organization-col", help="Explicit organization column name", default=None)
    args = parser.parse_args()

    initialize()

    sheet = args.sheet
    if isinstance(sheet, str) and sheet.isdigit():
        sheet = int(sheet)

    try:
        df = pd.read_excel(args.excel, sheet_name=sheet)
    except TypeError:
        df = pd.read_excel(args.excel, sheet_name=sheet, engine="openpyxl")

    df = df.fillna("")

    columns = [str(c) for c in df.columns.tolist()]
    topic_col = args.topic_col or _detect_column(columns, ["topic", "主題", "議題", "題目"])
    name_col = args.name_col or _detect_column(columns, ["name", "姓名"])
    title_col = args.title_col or _detect_column(columns, ["title", "職稱", "頭銜"])
    org_col = args.organization_col or _detect_column(
        columns,
        ["organization", "company", "單位", "服務單位", "機構", "組織", "affiliation", "部門"],
    )

    if not name_col:
        raise SystemExit("未能自動辨識姓名欄位，請使用 --name-col 指定。")

    rows: list[SignInRow] = []
    for _, row in df.iterrows():
        name_val = _strip_value(row.get(name_col))
        if not name_val:
            continue
        topic_val = _strip_value(row.get(topic_col)) if topic_col else ""
        title_val = _strip_value(row.get(title_col)) if title_col else ""
        org_val = _strip_value(row.get(org_col)) if org_col else ""
        rows.append(
            SignInRow(
                topic=topic_val,
                name=name_val,
                title=title_val,
                organization=org_val,
            )
        )

    if not rows:
        raise SystemExit("Excel 檔案中找不到任何有效的講者資料。")

    plan_name = args.plan_name or args.event_name or args.excel.stem
    event_name = args.event_name or plan_name
    date_display = format_date(args.date, sep="/") if args.date else None

    safe_event_name = _safe_filename_component(event_name)
    out_path = args.out or (OUTPUT_DIR / f"講師簽到表_{safe_event_name}.docx")

    context = SignInDocumentContext(
        plan_name=plan_name,
        event_name=event_name,
        date_display=date_display,
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
