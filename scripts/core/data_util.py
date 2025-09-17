from __future__ import annotations
import json
import re
import importlib
from pathlib import Path
from typing import Any, Dict, List, Optional

from .bootstrap import DATA_DIR

DEFAULT_DATA_DIR = DATA_DIR
DEFAULT_SHARED_JSON = DEFAULT_DATA_DIR / "shared" / "program_data.json"

# relaxed JSON loader (supports BOM, trailing commas)
def read_json_relaxed(p: Path) -> Any:
    s = p.read_text(encoding="utf-8")
    if s and s[0] == "\ufeff":
        s = s.lstrip("\ufeff")
    s = re.sub(r",\s*(?=[}\]])", "", s)
    return json.loads(s)


def load_programs(path: Optional[Path] = None) -> List[Dict[str, Any]]:
    path = Path(path) if path else DEFAULT_SHARED_JSON
    if not path.exists():
        return []
    try:
        data = read_json_relaxed(path)
    except Exception:
        return []
    if isinstance(data, dict):
        return [data]
    if isinstance(data, list):
        return data
    return []


def load_program_by_id(
    program_id: Optional[Any],
    *,
    path: Optional[Path] = None,
    fallback_to_first: bool = True,
) -> Dict[str, Any]:
    """Return a program dict matching ``program_id``.

    Parameters
    ----------
    program_id:
        The program ``id`` to look up. Accepts ``None`` (when ``fallback_to_first``
        is ``True``), ``int`` or ``str``.
    path:
        Optional path to a program_data.json file. Defaults to
        ``data/shared/program_data.json``.
    fallback_to_first:
        When ``True`` (default) and ``program_id`` is ``None`` or no match is
        found, the first program in the file is returned. When ``False`` a
        ``LookupError`` is raised instead.

    Raises
    ------
    FileNotFoundError
        If the JSON file does not exist.
    ValueError
        If the file is empty or contains no program entries.
    LookupError
        If ``program_id`` is provided but no matching program is found and
        ``fallback_to_first`` is ``False``.
    """

    path = Path(path) if path else DEFAULT_SHARED_JSON
    if not path.exists():
        raise FileNotFoundError(path)

    programs = load_programs(path)
    if not programs:
        raise ValueError("No program entries found in {}".format(path))

    if program_id is None:
        if fallback_to_first:
            return programs[0]
        raise LookupError("Program id is required")

    target = str(program_id).strip()
    for program in programs:
        pid = program.get("id")
        if pid is None:
            continue
        if str(pid).strip() == target:
            return program

    if fallback_to_first:
        return programs[0]
    raise LookupError("Program id {} not found in {}".format(program_id, path))


def find_data_file_by_id(data_dir: Path, target_id: str) -> Optional[Path]:
    target = str(target_id).strip()
    # JSON search
    for p in sorted(data_dir.rglob("*.json")):
        if p.resolve() == DEFAULT_SHARED_JSON.resolve():
            continue
        try:
            d = read_json_relaxed(p)
        except Exception:
            continue
        if isinstance(d, dict):
            # keys
            if any(str(k).strip() == target for k in d.keys()):
                return p
            if str(d.get("id", "")).strip() == target:
                return p
            for v in d.values():
                if isinstance(v, list):
                    for item in v:
                        if isinstance(item, dict) and str(item.get("id", "")).strip() == target:
                            return p
        elif isinstance(d, list):
            for item in d:
                try:
                    if isinstance(item, dict) and str(item.get("id", "")).strip() == target:
                        return p
                except Exception:
                    continue
    # Excel search (.xlsx / .xls)
    try:
        openpyxl = importlib.import_module("openpyxl")
    except ModuleNotFoundError:
        openpyxl = None
    try:
        xlrd = importlib.import_module("xlrd")
    except ModuleNotFoundError:
        xlrd = None

    for p in sorted(data_dir.rglob("*.xlsx")) + sorted(data_dir.rglob("*.xls")):
        if p.resolve() == DEFAULT_SHARED_JSON.resolve():
            continue
        ext = p.suffix.lower()
        if ext == ".xlsx" and openpyxl:
            try:
                wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
            except Exception:
                continue
            for sname in wb.sheetnames:
                ws = wb[sname]
                rows = list(ws.iter_rows(values_only=True))
                if not rows:
                    continue
                headers = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
                id_idx = [i for i, h in enumerate(headers) if "id" in h and h != ""]
                if not id_idx:
                    continue
                for row in rows[1:]:
                    for idx in id_idx:
                        if idx >= len(row):
                            continue
                        val = row[idx]
                        if val is None:
                            continue
                        if str(val).strip() == target:
                            wb.close()
                            return p
            wb.close()
        elif ext == ".xls" and xlrd:
            try:
                book = xlrd.open_workbook(str(p))
            except Exception:
                continue
            for sname in book.sheet_names():
                sh = book.sheet_by_name(sname)
                if sh.nrows == 0:
                    continue
                headers = [str(sh.cell_value(0, c)).strip().lower() for c in range(sh.ncols)]
                id_idx = [i for i, h in enumerate(headers) if "id" in h and h != ""]
                if not id_idx:
                    continue
                for r in range(1, sh.nrows):
                    for idx in id_idx:
                        try:
                            val = sh.cell_value(r, idx)
                        except Exception:
                            continue
                        if val is None or str(val).strip() == "":
                            continue
                        if str(val).strip() == target:
                            return p
    return None


def load_records(path: Path, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
    ext = path.suffix.lower()
    if ext == ".json":
        with path.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
        if isinstance(data, dict):
            if all(isinstance(v, dict) for v in data.values()):
                out = []
                for k, v in data.items():
                    rec = {str(kk).strip().lower(): vv for kk, vv in v.items()}
                    if "id" not in rec:
                        rec["id"] = str(k)
                    out.append(rec)
                return out
            data = [data]
        return [{str(k).strip().lower(): v for k, v in r.items()} for r in data]

    if ext in {".xlsx", ".xls"}:
        if ext == ".xlsx":
            openpyxl = importlib.import_module("openpyxl")
            wb = openpyxl.load_workbook(path, data_only=True)
            if sheet_name:
                if sheet_name not in wb.sheetnames:
                    raise ValueError("Sheet '{}' not found in {} (available: {})".format(sheet_name, path, wb.sheetnames))
                ws = wb[sheet_name]
            else:
                ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                return []
            headers = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
            recs: List[Dict[str, Any]] = []
            for row in rows[1:]:
                r = {}
                for i, val in enumerate(row):
                    header = headers[i] if i < len(headers) else "col_{}".format(i)
                    r[header] = val
                recs.append(r)
            return recs
        else:
            xlrd = importlib.import_module("xlrd")
            book = xlrd.open_workbook(str(path))
            if sheet_name:
                if sheet_name not in book.sheet_names():
                    raise ValueError("Sheet '{}' not found in {} (available: {})".format(sheet_name, path, book.sheet_names()))
                sh = book.sheet_by_name(sheet_name)
            else:
                sh = book.sheet_by_index(0)
            if sh.nrows == 0:
                return []
            headers = [str(sh.cell_value(0, c)).strip().lower() for c in range(sh.ncols)]
            recs = []
            for r in range(1, sh.nrows):
                rowvals = [sh.cell_value(r, c) for c in range(sh.ncols)]
                rec = {}
                for i, val in enumerate(rowvals):
                    header = headers[i] if i < len(headers) else "col_{}".format(i)
                    rec[header] = val
                recs.append(rec)
            return recs
    raise ValueError("Unsupported file extension: {}".format(ext))


def load_all_records_from_dir(data_dir: Path, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
    out = []
    for p in sorted(Path(data_dir).rglob("*")):
        if not p.is_file():
            continue
        if p.resolve() == DEFAULT_SHARED_JSON.resolve():
            continue
        if "shared" in p.parts:
            continue
        if p.suffix.lower() in {".json", ".xlsx", ".xls"}:
            try:
                recs = load_records(p, sheet_name=sheet_name)
                out.extend(recs)
            except Exception:
                continue
    return out


def record_matches_program(record: Dict[str, Any], program: Dict[str, Any]) -> bool:
    if not program:
        return False
    pid = str(program.get("id", "")).strip().lower()
    pname = str(program.get("planName", "") or program.get("plan_name", "")).strip().lower()
    candidate_keys = (
        "planid",
        "plan_id",
        "activity_id",
        "activityid",
        "program_id",
        "programid",
        "id",
        "planname",
        "plan_name",
        "program",
        "plan",
    )
    for k in candidate_keys:
        v = record.get(k)
        if v is None:
            continue
        vs = str(v).strip().lower()
        if vs == pid or vs == pname or (pname and pname in vs):
            return True
        if isinstance(v, (list, tuple)):
            for it in v:
                if str(it).strip().lower() in (pid, pname):
                    return True
        if isinstance(v, str) and "," in v:
            for it in [x.strip().lower() for x in v.split(",") if x.strip()]:
                if it in (pid, pname):
                    return True
    return False
