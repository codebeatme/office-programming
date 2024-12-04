# 读取 Excel 文件 Style.xlsx
from openpyxl import load_workbook
wb = load_workbook('Style.xlsx')

from openpyxl.styles.named_styles import NamedStyle

# ERROR 工作簿已经包含名称为 Normal 的命名格式
wb.add_named_style(NamedStyle('Normal'))
