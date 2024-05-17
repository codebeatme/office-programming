# 請將命令列跳躍至 1904.xlsx 所在的目錄，然後執行此腳本檔案
from openpyxl import load_workbook
workbook = load_workbook('1904.xlsx')

# 取得活頁簿的日期系統
print(workbook.excel_base_date)