# 讀取 Excel 檔案 Style.xlsx
from openpyxl import load_workbook
wb = load_workbook('Style.xlsx')

from openpyxl.styles.named_styles import NamedStyle

# ERROR 活頁簿已經包含名稱為 Normal 的命名格式
wb.add_named_style(NamedStyle('Normal'))
