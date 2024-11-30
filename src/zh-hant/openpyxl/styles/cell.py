# 讀取 Excel 檔案 Style.xlsx 中的工作表 Cell
from openpyxl import load_workbook
wb = load_workbook('Style.xlsx')
ws = wb['Cell']

from openpyxl.styles import Font, Border, PatternFill, Alignment, Color, Side

# 設定字型 Tahoma，15 磅，粗體，綠色，飽和度 50%
ws['A1'].font = Font('Tahoma', 15, True, color=Color('00FF00', tint=0.5))

# 顯示藍色的雙實線對角線
ws['B2'].border = Border(diagonal=Side('double', color='0000FF'), diagonalUp=True, diagonalDown=True)

# 圖樣 lightDown，前景藍色，背景紅色
ws['C3'].fill = PatternFill('lightDown', Color('0000FF'), Color('FF0000'))

# 顯示在右上角，在需要時縮小文字
ws['D4'].alignment = Alignment('right', 'top', shrinkToFit=True)

wb.save('Cell.xlsx')

# 可以修改字型色彩
ws['A1'].font.color.rgb = 'FFFF00'
# ERROR 不能直接修改具體格式
ws['A1'].font.b = True