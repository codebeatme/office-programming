# 讀取 Excel 檔案 Style.xlsx 中的工作表 Cell
from openpyxl import load_workbook
wb = load_workbook('Style.xlsx')
ws = wb['Cell']

from openpyxl.styles import Font, Border, PatternFill, Alignment, Side

# 為第一列設定字型 Arial，雙實線底線
ws.row_dimensions[1].font = Font('Arial', u='double')
# 為第二欄設定紅色底邊
ws.column_dimensions['B'].border = Border(bottom=Side('dashed', 'FF0000'))
# 為第四欄填入綠色
ws.column_dimensions['D'].fill = PatternFill('solid', '00FF00')
# 為第四列設定左下角對齊
ws.row_dimensions[4].alignment = Alignment('left', 'bottom')

wb.save('RC.xlsx')