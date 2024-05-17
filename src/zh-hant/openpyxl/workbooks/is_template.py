# 請將命令列跳躍至 Hello.xltx 所在的目錄，然後執行此腳本檔案
from openpyxl import load_workbook
workbook = load_workbook('Hello.xltx')

# 活頁簿是否為範本？
print('template' in workbook.mime_type)