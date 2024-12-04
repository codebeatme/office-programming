# 讀取 Excel 檔案 Style.xlsx
from openpyxl import load_workbook
wb = load_workbook('Style.xlsx')

# 顯示 Excel 活頁簿中的所有命名格式的名稱
print(wb.style_names)
print(wb.named_styles)