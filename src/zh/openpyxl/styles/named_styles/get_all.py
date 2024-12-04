# 读取 Excel 文件 Style.xlsx
from openpyxl import load_workbook
wb = load_workbook('Style.xlsx')

# 显示 Excel 工作簿中的所有命名格式的名称
print(wb.style_names)
print(wb.named_styles)