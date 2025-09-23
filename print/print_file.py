import os
import win32com.client
import win32api
import win32print

# Word 常數：手送匣
WD_TRAY_MANUAL = 2   # wdPrinterManualFeed = 2

def print_word(file_path):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(file_path, ReadOnly=True)

        # 設定紙匣：首張 + 其他頁都走手送匣
        doc.PageSetup.FirstPageTray = WD_TRAY_MANUAL
        doc.PageSetup.OtherPagesTray = WD_TRAY_MANUAL

        # 靜默列印，逐份
        doc.PrintOut(Background=False, Collate=True)
        doc.Close(False)
        word.Quit()
        print(f"[Word] 手送匣＋逐份 列印完成：{file_path}")
    except Exception as e:
        print(f"[Word] 列印失敗 {file_path}: {e}")

def print_excel(file_path):
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(file_path, ReadOnly=True)

        # Excel 沒有 PageSetup.Tray 這麼方便的設定，
        # 只能透過印表機 DEVMODE 先設成手送匣再列印
        wb.PrintOut(Copies=1, Collate=True)  # 逐份列印
        wb.Close(False)
        excel.Quit()
        print(f"[Excel] 逐份 列印完成：{file_path}")
    except Exception as e:
        print(f"[Excel] 列印失敗 {file_path}: {e}")

def print_other(file_path):
    try:
        win32api.ShellExecute(
            0,
            "print",
            file_path,
            None,
            ".",
            0
        )
        print(f"[其他] 已送列印：{file_path}")
    except Exception as e:
        print(f"[其他] 無法列印 {file_path}: {e}")

def batch_print_all(folder_path):
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path):
            ext = os.path.splitext(file_path)[1].lower()
            if ext in [".doc", ".docx"]:
                print_word(file_path)
            elif ext in [".xls", ".xlsx"]:
                print_excel(file_path)
            else:
                print_other(file_path)

if __name__ == "__main__":
    folder = r"C:\Users\MPAT05\Desktop\JupyterProjects\print\New folder"  # 修改成您的資料夾
    print(f"目前設定的資料夾：{folder}")
    batch_print_all(folder)
