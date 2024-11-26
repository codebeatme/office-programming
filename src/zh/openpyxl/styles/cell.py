# 读取 Excel 文件 Style.xlsx 中的工作表 Cell
from openpyxl import load_workbook
wb = load_workbook('Style.xlsx')
ws = wb['Cell']

from openpyxl.styles import Font, Border, PatternFill, Alignment, Color, Side

# 设置字体 Tahoma，15 磅，粗体，绿色，饱和度 50%
ws['A1'].font = Font('Tahoma', 15, True, color=Color('00FF00', tint=0.5))

# 显示蓝色的双实线对角线
ws['B2'].border = Border(diagonal=Side('double', color='0000FF'), diagonalUp=True, diagonalDown=True)

# 图案 lightDown，前景蓝色，背景红色
ws['C3'].fill = PatternFill('lightDown', Color('0000FF'), Color('FF0000'))

# 显示在右上角，在需要时缩小文字
ws['D4'].alignment = Alignment('right', 'top', shrinkToFit=True)

wb.save('Cell.xlsx')

# 可以修改字体颜色
ws['A1'].font.color.rgb = 'FFFF00'
# ERROR 不能直接修改具体样式
ws['A1'].font.b = True