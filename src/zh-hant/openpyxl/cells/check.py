# 開啟 Data.xlsx 中的工作表 Values
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Values']

a1 = worksheet['A1']
# 嘗試轉換為有效的文字內容
a1.check_string(b'\0 is invalid')