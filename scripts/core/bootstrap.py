# scripts/core/bootstrap.py
import json
import csv
import os
import shutil
import platform
from pathlib import Path
from typing import Optional

# Project root
BASE_DIR = Path(__file__).resolve().parents[2]

# load config/paths.json
_config_path = BASE_DIR / "config" / "paths.json"
try:
    with _config_path.open("r", encoding="utf-8") as f:
        _PATHS = json.load(f)
except Exception:
    _PATHS = {}

# base folder handling: if BaseFolder is absolute use it, otherwise treat as relative to BASE_DIR
_raw_base = _PATHS.get("BaseFolder", "")
_base = Path(_raw_base) if isinstance(_raw_base, str) and Path(_raw_base).is_absolute() else (BASE_DIR / Path(_raw_base))

def _resolve(key: str, default: str) -> Path:
    """
    Resolve a path from _PATHS:
      - Replace ${BaseFolder}, expand ~ and env vars ($VAR or %VAR%)
      - If resulting path is absolute -> return as-is
      - Else -> return _base / relative_path
    """
    raw_val = _PATHS.get(key, default)
    if isinstance(raw_val, str):
        s = raw_val.replace("${BaseFolder}", str(_base))
        s = os.path.expanduser(os.path.expandvars(s))
        p = Path(s)
    else:
        p = Path(raw_val)

    return p if p.is_absolute() else (_base / p)

# Export commonly used paths
DATA_DIR = _resolve("Data", "data")
TEMPLATE_DIR = _resolve("Templates", "templates")
OUTPUT_DIR = _resolve("Output", "output")
PROGRAM_JSON = DATA_DIR / "program.json"

# Optional extras (if present in paths.json)
OUTPUT_LETTERS = _resolve("OutputLetters", str(OUTPUT_DIR / "letters"))
OUTPUT_REPORTS = _resolve("OutputReports", str(OUTPUT_DIR / "reports"))
ACTIVITIES_DATA = _resolve("ActivitiesData", str(DATA_DIR / "activities"))

def initialize():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def search_file(base_dir: Path, target_filename: str) -> Path:
    for path in base_dir.rglob("*"):
        if path.name == target_filename:
            return path
    raise FileNotFoundError(f"找不到檔案: {target_filename}")

def load_schema(filename: str) -> dict:
    path = search_file(BASE_DIR / "config" / "schema", filename)
    with open(path, encoding="utf-8") as f:
        return json.load(f)

def load_json_file(filename: str) -> dict:
    path = search_file(DATA_DIR, filename)
    with open(path, encoding="utf-8") as f:
        return json.load(f)

def load_csv_file(filename: str) -> list[dict]:
    path = search_file(DATA_DIR, filename)
    with open(path, newline='', encoding="utf-8") as f:
        return list(csv.DictReader(f))

def merge_schema(schema: dict, data_list: list) -> list:
    result = []
    for data in data_list:
        merged = {key: data.get(key, schema[key]) for key in schema}
        result.append(merged)
    return result

# -----------------------------
# Chrome binary resolution logic
# Exports CHROME_BIN (str) or None if not found
# Priority:
#   1. Environment variable CHROME_BIN or CHROME_PATH
#   2. "Chrome" value in config/paths.json (supports ${BaseFolder}, env vars, ~)
#   3. common PATH names via shutil.which
#   4. common platform install locations
# -----------------------------
def _expand_value(val: Optional[str]) -> Optional[str]:
    if not isinstance(val, str) or not val:
        return None
    s = val.replace("${BaseFolder}", str(_base))
    s = os.path.expanduser(os.path.expandvars(s))
    return s

def _find_chrome_from_config() -> Optional[str]:
    raw = _PATHS.get("Chrome")
    if raw:
        cand = _expand_value(raw)
        if cand and Path(cand).exists():
            return str(Path(cand))
    return None

def _find_chrome_on_path() -> Optional[str]:
    candidates = [
        "google-chrome", "google-chrome-stable", "chrome", "chromium", "chromium-browser", "msedge",
        "chrome.exe", "msedge.exe"
    ]
    for name in candidates:
        p = shutil.which(name)
        if p:
            return p
    return None

def _find_chrome_by_common_locations() -> Optional[str]:
    system = platform.system()
    if system == "Windows":
        win_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        ]
        for p in win_paths:
            if Path(p).exists():
                return str(Path(p))
    elif system == "Darwin":
        mac_paths = [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Chromium.app/Contents/MacOS/Chromium",
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
        ]
        for p in mac_paths:
            if Path(p).exists():
                return p
    else:
        linux_paths = [
            "/usr/bin/google-chrome",
            "/usr/bin/google-chrome-stable",
            "/usr/bin/chromium-browser",
            "/usr/bin/chromium",
            "/snap/bin/chromium",
        ]
        for p in linux_paths:
            if Path(p).exists():
                return p
    return None

# Final resolution
_env_override = os.environ.get("CHROME_BIN") or os.environ.get("CHROME_PATH")
CHROME_BIN: Optional[str] = None
if _env_override and Path(_env_override).exists():
    CHROME_BIN = str(Path(_env_override))
else:
    CHROME_BIN = _find_chrome_from_config() or _find_chrome_on_path() or _find_chrome_by_common_locations()

# CHROME_BIN may be None if not found — scripts should handle that case
