# 读取 Excel 文件 Style.xlsx 中的工作表 Cell
from openpyxl import load_workbook
wb = load_workbook('Style.xlsx')
ws = wb['Cell']

for s in wb.named_styles:
    print(f'"{s}"', type(s))


for s in wb.style_names:
    print(f'"{s}"', type(s))

wb.add_named_style()