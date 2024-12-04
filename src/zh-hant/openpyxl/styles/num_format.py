# 讀取 Excel 檔案 Style.xlsx 中的工作表 Cell
from openpyxl import load_workbook
wb = load_workbook('Style.xlsx')
ws = wb['Cell']

# 使用千位分隔符，如果是負數，則顯示為紅色
ws['E5'].number_format = '#,##0;[RED]-#,##0'

wb.save('Format.xlsx')
