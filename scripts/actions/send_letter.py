# -*- coding: utf-8 -*-
from pathlib import Path
import json
from docxtpl import DocxTemplate

# 這三個 import 僅在 Windows + 安裝 Outlook 時可用
import win32com.client as win32
import pythoncom

from scripts.core.bootstrap import (
    initialize, DATA_DIR, TEMPLATE_DIR, OUTPUT_DIR,
    search_file
)

def _render_docx(template_path: Path, context: dict, out_path: Path):
    out_path.parent.mkdir(parents=True, exist_ok=True)
    tpl = DocxTemplate(str(template_path))
    tpl.render(context)
    tpl.save(str(out_path))
    return out_path

def _maybe_attach_pdf(pdf_dir: str | Path | None, person: dict, pdf_map_by: str) -> list[Path]:
    if not pdf_dir:
        return []
    key = person.get(pdf_map_by)
    if not key:
        return []
    cand = Path(pdf_dir) / f"{key}.pdf"
    return [cand] if cand.exists() else []

def _ensure_outlook():
    """取得或啟動 Outlook Application 物件。"""
    pythoncom.CoInitialize()
    try:
        app = win32.gencache.EnsureDispatch("Outlook.Application")
    except Exception:
        app = win32.Dispatch("Outlook.Application")
    return app

def _create_outlook_mail(app, to_email: str, subject: str, html_body: str, attachments: list[Path]):
    mail = app.CreateItem(0)  # 0 = olMailItem
    mail.To = to_email
    mail.Subject = subject
    mail.HTMLBody = html_body
    for p in attachments or []:
        if p.exists():
            mail.Attachments.Add(str(p))
    return mail

def send_letter(
    filter_by: str,                  # 在 event_contacts.json 中用來篩選的欄位，如 "group"
    group_value: str,                # 欲篩選的值，如 "influencer"
    data_file: str,                  # JSON 檔名，如 "activities_contacts.json"（會在 data/** 下遞迴找）
    template_name: str,              # Word 模板（templates/letters/ 下）
    send_mode: str = "draft",        # "draft" 存到草稿匣；"send" 直接寄出；"display" 打開編輯視窗
    pdf_dir: str | None = None,      # 附件資料夾（可選）
    pdf_map_by: str = "name",        # PDF 檔名對應欄位（name 或 id）
    subject_fmt: str = "邀請函－{name}",
    html_fmt: str = "<p>您好 {name}：</p><p>邀請函請見附件與 Word 檔。</p>",
    attach_generated_docx: bool = True  # 是否把產生的 docx 也一併附上
):
    """
    從 data/**/data_file 讀 event contacts，依 filter_by == group_value 篩選，
    使用 templates/letters/template_name 產生個人化 Word，並透過 Outlook 建信：
      - send_mode="draft"：存到草稿匣
      - send_mode="display"：開啟視窗讓你再確認
      - send_mode="send"：直接寄出
    """
    initialize()

    # 1) 找資料檔
    contacts_path = search_file(DATA_DIR, data_file)
    with open(contacts_path, "r", encoding="utf-8") as f:
        contacts = json.load(f)

    # 2) 篩選
    people = [p for p in contacts if p.get(filter_by) == group_value]
    if not people:
        print(f"[警告] 在 {contacts_path} 內找不到 {filter_by} = {group_value} 的資料")
        return

    # 3) 模板路徑
    tpl_path = TEMPLATE_DIR / "letters" / template_name
    if not tpl_path.exists():
        raise FileNotFoundError(f"找不到模板：{tpl_path}")

    # 4) 輸出資料夾
    letters_dir = OUTPUT_DIR / "letters"
    letters_dir.mkdir(parents=True, exist_ok=True)

    # 5) 取得 Outlook
    app = _ensure_outlook()
    ns = app.GetNamespace("MAPI")  # 讓 Outlook 綁定 MAPI（有些環境需要）

    # 6) 逐筆處理
    for person in people:
        # 6-1 產 DOCX
        safe_name = str(person.get("name", "unknown")).replace("/", "_").replace("\\", "_").strip()
        docx_out = letters_dir / f"{safe_name}_{group_value}.docx"
        _render_docx(tpl_path, person, docx_out)
        print(f"✅ 產出 Word：{docx_out}")

        # 6-2 附件（PDF + 產生的 DOCX（可選））
        attachments = _maybe_attach_pdf(pdf_dir, person, pdf_map_by)
        if attach_generated_docx:
            attachments.append(docx_out)

        # 6-3 準備郵件內容
        mapping = {k: str(v) for k, v in person.items()}
        subject = subject_fmt.format(**mapping)
        html_body = html_fmt.format(**mapping)

        to_email = person.get("email")
        if not to_email:
            print(f"[跳過] {safe_name} 缺少 email 欄位")
            continue

        mail = _create_outlook_mail(app, to_email, subject, html_body, attachments)

        # 6-4 送出、顯示、或存草稿
        mode = send_mode.lower()
        if mode == "send":
            mail.Send()
            print(f"📤 已寄出：{to_email}")
        elif mode == "display":
            mail.Display(True)  # True = 模式化視窗（阻塞）
            print(f"📝 已開啟編輯視窗：{to_email}")
        else:
            # 存到 Drafts（草稿匣）
            mail.Save()
            print(f"✉️ 草稿已存：{safe_name} -> Drafts")

# CLI 入口（可直接用 -m 執行）
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--filter-by", required=True, help="用於篩選的欄位名（如 group/role/type）")
    parser.add_argument("--group-value", required=True, help="篩選值（如 influencer/participant 等）")
    parser.add_argument("--data-file", required=True, help="資料檔名（如 activities_contacts.json）")
    parser.add_argument("--template", required=True, help="Word 模板檔名（放在 templates/letters/）")
    parser.add_argument("--send-mode", default="draft", choices=["draft", "display", "send"])
    parser.add_argument("--pdf-dir", default=None)
    parser.add_argument("--pdf-map-by", default="name")
    parser.add_argument("--subject-fmt", default="邀請函－{name}")
    parser.add_argument("--html-fmt", default="<p>您好 {name}：</p><p>邀請函請見附件與 Word 檔。</p>")
    parser.add_argument("--attach-generated-docx", action="store_true")
    args = parser.parse_args()

    send_letter(
        filter_by=args.filter_by,
        group_value=args.group_value,
        data_file=args.data_file,
        template_name=args.template,
        send_mode=args.send_mode,
        pdf_dir=args.pdf_dir,
        pdf_map_by=args.pdf_map_by,
        subject_fmt=args.subject_fmt,
        html_fmt=args.html_fmt,
        attach_generated_docx=args.attach_generated_docx
    )
