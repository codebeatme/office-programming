# 讀取 Excel 檔案 Style.xlsx 中的工作表 Font
from openpyxl import load_workbook
wb = load_workbook('Style.xlsx')
ws = wb['Font']

from openpyxl.styles import Font, Color

a1 = ws['A1']
# 設定字型 Tahoma，15 磅，粗體，綠色，飽和度 50%
a1.font = Font('Tahoma', 15, True, color=Color('00FF00', tint=0.5))

b1 = ws['B1']
# 取得字型
print(f'B1 目前字型：{b1.font}')

wb.save('SetFont.xlsx')

# 可以修改字型色彩
b1.font.color.rgb = 'FFFF00'
# ERROR 不能直接修改具體樣式
b1.font.b = True