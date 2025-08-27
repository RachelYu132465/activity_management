# scripts/actions/font_test_render.py
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Font load + render test for nameplate generator.
Saves output to OUTPUT_DIR / "font_test_render.png".
"""

from __future__ import annotations
import sys
import platform
from pathlib import Path
from typing import List, Tuple
from PIL import Image, ImageDraw, ImageFont

# project root detection (same as your other scripts)
ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# try import bootstrap constants if available, else fallback to local output
try:
    from scripts.core.bootstrap import OUTPUT_DIR, DATA_DIR, initialize
    try:
        initialize()
    except Exception:
        pass
except Exception:
    OUTPUT_DIR = ROOT / "output"
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Candidate lists based on your system listing
MSJH_BOLD_CANDIDATES = [
    r"C:\Windows\Fonts\msjhbd.ttc",
    r"C:\Windows\Fonts\msjhbd.ttf",
    r"C:\Windows\Fonts\msjhl.ttc",
    r"C:\Windows\Fonts\msjh.ttc",
    r"C:\Windows\Fonts\msjh.ttf",
    # project-local fallback (if you copy fonts into scripts/fonts/)
    str(ROOT / "scripts" / "fonts" / "msjhbd.ttf"),
]

MSJH_REGULAR_CANDIDATES = [
    r"C:\Windows\Fonts\msjh.ttc",
    r"C:\Windows\Fonts\msjh.ttf",
    r"C:\Windows\Fonts\msyh.ttc",
    str(ROOT / "scripts" / "fonts" / "msjh.ttf"),
]

BKAI_CANDIDATES = [
    r"C:\Windows\Fonts\kaiu.ttf",
    r"C:\Windows\Fonts\STKAITI.TTF",
    r"C:\Windows\Fonts\STXINGKA.TTF",
    r"C:\Windows\Fonts\NotoSansTC-VF.ttf",
    str(ROOT / "scripts" / "fonts" / "DFKai-SB.ttf"),
]

# also try some common family names
COMMON_NAMES = ["DejaVuSans.ttf", "Arial.ttf", "NotoSansTC-VF.ttf", "msjhbd.ttf", "msjh.ttf", "kaiu.ttf"]

def try_load_font(candidates: List[str], size: int, label: str) -> Tuple[ImageFont.FreeTypeFont, str|None]:
    """Try load fonts from candidates, then common names. Return (font_obj, source)."""
    for p in candidates:
        if not p:
            continue
        try:
            pth = Path(p)
            if pth.exists():
                f = ImageFont.truetype(str(pth), size)
                print(f"[OK] {label}: loaded from path -> {pth}")
                return f, str(pth)
        except Exception as e:
            print(f"[FAIL] {label}: tried path {p!r} -> {e}")
    # try common names
    for nm in COMMON_NAMES:
        try:
            f = ImageFont.truetype(nm, size)
            print(f"[OK] {label}: loaded by system name -> {nm}")
            return f, nm
        except Exception as e:
            print(f"[FAIL] {label}: tried system name {nm!r} -> {e}")
    # fallback
    print(f"[WARN] {label}: no TTF found, using ImageFont.load_default() (bitmap, not scalable).")
    return ImageFont.load_default(), None

def text_bbox_size(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont) -> Tuple[int,int]:
    if not text:
        return 0,0
    bbox = draw.textbbox((0,0), text, font=font)
    w = bbox[2] - bbox[0]
    h = bbox[3] - bbox[1]
    return int(w), int(h)

def main():
    OUT = OUTPUT_DIR / "font_test_render.png"
    W, H = 1200, 400

    # choose sizes to test (large for name, medium for title/org)
    NAME_SIZE = 160
    TITLE_SIZE = 48
    ORG_SIZE = 44

    print("=== Attempting to load fonts ===")
    name_font, name_src = try_load_font(MSJH_BOLD_CANDIDATES, NAME_SIZE, "Name-Bold")
    title_font, title_src = try_load_font(MSJH_REGULAR_CANDIDATES, TITLE_SIZE, "Title-Regular")
    org_font, org_src = try_load_font(BKAI_CANDIDATES, ORG_SIZE, "Org-BKAI")

    print("\nFont selection summary:")
    print(" Name font:", name_src)
    print(" Title font:", title_src)
    print(" Org font:", org_src)
    print("=================================\n")

    # Draw sample lines
    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    sample_org = "單位：國立測試大學 學術與研究處"
    sample_name = "林世嘉"
    sample_title = "特聘教授 / 跨領域研究"

    # draw org top-left
    ox, oy = 40, 20
    draw.text((ox, oy), sample_org, font=org_font, fill="black")
    w_o, h_o = text_bbox_size(draw, sample_org, org_font)
    print(f"[bbox] org: {w_o}x{h_o}")

    # draw big name left-biased, vertically centered
    x_name = 80
    w_name, h_name = text_bbox_size(draw, sample_name, name_font)
    # if name_font is bitmap fallback, warn
    if name_src is None:
        print("[WARN] Name font is fallback bitmap (small). You must install / provide a TTF to get big text.")
    y_name = (H // 2) - (h_name // 2)
    draw.text((x_name, y_name), sample_name, font=name_font, fill="black")
    print(f"[bbox] name: {w_name}x{h_name}")

    # draw title on right, right-aligned
    right_x = W - 40
    # compute bbox of title
    w_title, h_title = text_bbox_size(draw, sample_title, title_font)
    draw.text((right_x - w_title, (H // 2) - (h_title // 2)), sample_title, font=title_font, fill="black")
    print(f"[bbox] title: {w_title}x{h_title}")

    # save
    OUT.parent.mkdir(parents=True, exist_ok=True)
    img.save(OUT)
    print(f"\nSaved render to: {OUT}")
    print("請開啟該檔案看實際效果；若仍是極小字或跑到角落，代表載入的 font_src = None（bitmap fallback）。")
    print("若 fallback 發生，請把你要用的 TTF 複製到 scripts/fonts/ 並重新執行此測試，或回報下列資訊：")
    print(" - 你看到的輸出檔路徑\n - 上面列出的 [FAIL] / [OK] 訊息")

if __name__ == "__main__":
    main()
