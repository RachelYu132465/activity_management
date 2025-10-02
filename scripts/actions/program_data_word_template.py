"""Generate a Word document listing ``program_data`` placeholders."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Iterator, Sequence

try:
    from docx import Document
except ImportError as exc:  # pragma: no cover - dependency hint for CLI use
    raise SystemExit("請先安裝 python-docx：pip install python-docx") from exc

from scripts.core.bootstrap import DATA_DIR, OUTPUT_DIR


def load_programs(path: Path) -> Sequence[dict[str, Any]]:
    """Load program records from ``path`` and normalize to a sequence."""

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:  # pragma: no cover - guard rail for CLI
        raise SystemExit(f"找不到 program_data：{path}") from exc

    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        return [payload]
    raise SystemExit("program_data.json 內容必須是 JSON 物件或陣列")


def iter_paths(node: Any, prefix: str) -> Iterator[tuple[str, Any]]:
    """Yield dotted paths and scalar values from ``node``."""

    if isinstance(node, dict):
        for key, value in node.items():
            child_prefix = f"{prefix}.{key}" if prefix else key
            yield from iter_paths(value, child_prefix)
    elif isinstance(node, list):
        for idx, value in enumerate(node):
            child_prefix = f"{prefix}[{idx}]"
            yield from iter_paths(value, child_prefix)
    else:
        yield (prefix, node)


def format_value(value: Any) -> str:
    """Convert a scalar value into a display string for the Word table."""

    if value is None:
        return ""
    if isinstance(value, bool):
        return "True" if value else "False"
    return str(value)


def create_document(
    programs: Sequence[dict[str, Any]], *, include_values: bool, output_path: Path
) -> Path:
    """Build the Word document and save it to ``output_path``."""

    doc = Document()
    doc.add_heading("program_data Placeholders", level=1)
    doc.add_paragraph(
        "此文件列出可在 Word 模板中使用的 {{ program_data... }} 變數。"
        "您可以將這些變數貼到模板中，以便後續自動填入資料。"
    )

    for program in programs:
        program_id = program.get("id", "?")
        event_names = program.get("eventNames") or []
        if isinstance(event_names, list) and event_names:
            event_name = str(event_names[0])
        else:
            event_name = ""

        heading = doc.add_heading(level=2)
        heading.text = f"Program ID {program_id} {event_name}".strip()

        columns = 2 if include_values else 1
        table = doc.add_table(rows=1, cols=columns)
        header_cells = table.rows[0].cells
        header_cells[0].text = "Placeholder"
        if include_values:
            header_cells[1].text = "目前資料"

        for path, value in iter_paths(program, "program_data"):
            row_cells = table.add_row().cells
            row_cells[0].text = f"{{{{ {path} }}}}"
            if include_values:
                row_cells[1].text = format_value(value)

        doc.add_paragraph("")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="產生列出 program_data 變數的 Word 範本表")
    parser.add_argument(
        "--data",
        type=Path,
        default=DATA_DIR / "shared" / "program_data.json",
        help="program_data.json 的路徑 (預設: data/shared/program_data.json)",
    )
    parser.add_argument(
        "--program-id",
        type=int,
        action="append",
        dest="program_ids",
        help="限定輸出的 program id，可重複指定多個。未指定時輸出所有。",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=OUTPUT_DIR / "program_data_placeholders.docx",
        help="輸出的 Word 檔案路徑 (預設: output/program_data_placeholders.docx)",
    )
    parser.add_argument(
        "--no-values",
        action="store_true",
        help="只列出變數名稱，不顯示目前資料值。",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    programs = load_programs(args.data)

    if args.program_ids:
        wanted = set(args.program_ids)
        programs = [p for p in programs if p.get("id") in wanted]
        if not programs:
            raise SystemExit("找不到指定 id 的 program 資料")

    output_path = args.output
    include_values = not args.no_values
    create_document(programs, include_values=include_values, output_path=output_path)
    print(f"[OK] 已輸出：{output_path}")


if __name__ == "__main__":
    main()
