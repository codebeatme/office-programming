# 讀取 Excel 檔案 Style.xlsx 中的工作表 Cell
from openpyxl import load_workbook
wb = load_workbook('Style.xlsx')
ws = wb['Cell']

from openpyxl.styles.named_styles import NamedStyle

# 將儲存格 A1 的命名格式設定為預設的 Comma
ws['A1'].style = 'Comma'

# 將儲存格 B1 的命名格式設定為 MyStyle
b1 = ws['B1']
print(b1.style)
b1.style = NamedStyle('MyStyle')
print(b1.style)

wb.save('Cell.xlsx')
