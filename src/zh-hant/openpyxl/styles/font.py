from openpyxl.styles import Font, Color

# 粗體，斜體，單底線，刪除線
Font(b=True, i=True, u=Font.UNDERLINE_SINGLE, strike=True)
# 輪廓，陰影，20 磅
Font(outline=True, shadow=True, sz=20)
# 下標
Font(vertAlign='superscript')
# 紅色，飽和度 50%
Font(color=Color(rgb='FF0000', tint=0.5))
# 藍色
Font(color='0000FF')