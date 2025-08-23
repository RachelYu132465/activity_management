# render_with_chrome.py
# 作用：用 Jinja2 產生 HTML（templates/template.html + data/program.json）
# 然後用本機 Chrome/Edge headless 把 HTML 轉成 PDF（output/program.pdf）

import json
import os
import shutil
import subprocess
from pathlib import Path
from jinja2 import Environment, FileSystemLoader, select_autoescape

# --- 設定路徑（視你的專案結構調整） ---
TEMPLATES_DIR = Path("templates")
DATA_FILE = Path("data/program_data.json")
OUT_DIR = Path("output")
OUT_HTML = OUT_DIR / "program.html"
OUT_PDF = OUT_DIR / "program.pdf"

OUT_DIR.mkdir(parents=True, exist_ok=True)

# --- 讀資料 & render HTML ---
with DATA_FILE.open("r", encoding="utf8") as f:
    data = json.load(f)

env = Environment(loader=FileSystemLoader(str(TEMPLATES_DIR)),
                  autoescape=select_autoescape(["html", "xml"]))
tpl = env.get_template("template.html")
html = tpl.render(**data)

OUT_HTML.write_text(html, encoding="utf8")
print(f"HTML preview written to: {OUT_HTML.resolve()}")

# --- 自動尋找可用的 chrome/edge 可執行檔 ---
candidates = [
    # 常見 Chrome 路徑（Windows）
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    # 常見 Edge 路徑（Windows）
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    # 可被 shutil.which 找到的命令
    "chrome",
    "chromium",
    "msedge",
    "google-chrome"
]

chrome_path = None
for c in candidates:
    if Path(c).is_file():
        chrome_path = str(Path(c).resolve())
        break
    found = shutil.which(c)
    if found:
        chrome_path = found
        break

if not chrome_path:
    raise RuntimeError(
        "找不到 Chrome / Edge 可執行檔。請安裝 Chrome 或 Edge，或改用 WeasyPrint/wkhtmltopdf。"
    )

# --- 使用 headless chrome 產生 PDF ---
# 使用 Path.as_uri() 來取得正確的 file:///URI
html_uri = OUT_HTML.resolve().as_uri()
cmd = [
    chrome_path,
    "--headless",
    "--disable-gpu",
    # 可選: "--no-sandbox",
    f"--print-to-pdf={str(OUT_PDF.resolve())}",
    html_uri
]

print("Running:", " ".join(cmd))
subprocess.run(cmd, check=True)
print(f"PDF written to: {OUT_PDF.resolve()}")
