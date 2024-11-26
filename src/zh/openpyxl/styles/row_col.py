# 读取 Excel 文件 Style.xlsx 中的工作表 Cell
from openpyxl import load_workbook
wb = load_workbook('Style.xlsx')
ws = wb['Cell']

from openpyxl.styles import Font, Border, PatternFill, Alignment, Side

# 为第一行设置字体 Arial，双实线下划线
ws.row_dimensions[1].font = Font('Arial', u='double')
# 为第二列设置红色底边
ws.column_dimensions['B'].border = Border(bottom=Side('dashed', 'FF0000'))
# 为第四列填充绿色
ws.column_dimensions['D'].fill = PatternFill('solid', '00FF00')
# 为第四行设置左下角对齐
ws.row_dimensions[4].alignment = Alignment('left', 'bottom')

wb.save('RC.xlsx')