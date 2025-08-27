#!/usr/bin/env python3
# update_docx_fields.py
# Usage: python update_docx_fields.py "C:\path\to\your\test.docx" [--visible] [--restart-page-number]
# Requires Windows + Microsoft Word + pywin32

import sys
import os
from pathlib import Path

def update_docx_fields(docx_path: str, visible: bool = False, restart_page_number: bool = False) -> None:
    try:
        from win32com.client import Dispatch, constants
    except Exception as e:
        raise RuntimeError("pywin32 is required. Install with: pip install pywin32") from e

    # Normalize path and verify
    p = Path(docx_path)
    if not p.exists():
        raise FileNotFoundError(f"File not found: {docx_path}")

    word = Dispatch("Word.Application")
    word.Visible = bool(visible)
    # Open document (Read/Write)
    doc = word.Documents.Open(str(p.resolve()))

    try:
        # Update all fields (general)
        doc.Fields.Update()

        # Update every Table of Contents if present
        toc_count = doc.TablesOfContents.Count
        if toc_count > 0:
            for i in range(1, toc_count + 1):
                try:
                    doc.TablesOfContents(i).Update()
                except Exception:
                    # best-effort
                    pass

        # Optionally restart page numbering at section 2 (common: cover is first section)
        if restart_page_number:
            try:
                # If there's at least 2 sections, restart numbering at section 2; otherwise restart at section 1
                secs = doc.Sections.Count
                target_idx = 2 if secs >= 2 else 1
                sec = doc.Sections(target_idx)

                # Restart page numbers in the primary footer of that section
                sec.Footers(constants.wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
                sec.Footers(constants.wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
            except Exception as e:
                # not fatal; print warning and continue
                print("Warning: couldn't set restart page numbering:", e)

        # Save & close
        doc.Save()
    finally:
        doc.Close(False)
        word.Quit()

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Update Word fields and TOC via Word COM (Windows+Word required).")
    parser.add_argument("docx", help="Path to .docx file")
    parser.add_argument("--visible", action="store_true", help="Show Word UI while updating (for debugging)")
    parser.add_argument("--restart-page-number", action="store_true", help="Restart page numbers at section 2 (common for cover).")
    args = parser.parse_args()

    try:
        update_docx_fields(args.docx, visible=args.visible, restart_page_number=args.restart_page_number)
        print("Updated fields and TOC successfully:", args.docx)
    except Exception as e:
        print("Error:", e)
        sys.exit(1)
