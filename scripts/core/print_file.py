import os
import win32api

def batch_print_all(folder_path):
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path):
            try:
                print(f"列印：{file_path}")
                win32api.ShellExecute(
                    0,
                    "print",
                    file_path,
                    None,
                    ".",
                    0
                )
            except Exception as e:
                print(f"無法列印 {file_path}: {e}")

if __name__ == "__main__":
    folder = r"D:\mydesktop\print"  # 修改成您的資料夾
    batch_print_all(folder)
