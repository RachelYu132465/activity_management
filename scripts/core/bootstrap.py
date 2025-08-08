import json
import csv
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = BASE_DIR / "data"
TEMPLATE_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"

def initialize():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def search_file(base_dir, target_filename):
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
