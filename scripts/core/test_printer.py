import win32print

DC_BINS = 6
DC_BINNAMES = 12

# 先取得預設印表機
printer_name = win32print.GetDefaultPrinter()

# 開啟印表機，拿到 port 名稱
ph = win32print.OpenPrinter(printer_name)
printer_info = win32print.GetPrinter(ph, 2)   # level 2 info
port_name = printer_info["pPortName"]
win32print.ClosePrinter(ph)

# 查詢支援的紙匣
bins = win32print.DeviceCapabilities(printer_name, port_name, DC_BINS)
names = win32print.DeviceCapabilities(printer_name, port_name, DC_BINNAMES)

for code, name in zip(bins, names):
    print(code, name.strip())
