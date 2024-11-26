from openpyxl.styles import Font, Color

# 粗体，斜体，单下划线，删除线
Font(b=True, i=True, u=Font.UNDERLINE_SINGLE, strike=True)
# 轮廓，阴影，20 磅
Font(outline=True, shadow=True, sz=20)
# 下标
Font(vertAlign='superscript')
# 红色，饱和度 50%
Font(color=Color(rgb='FF0000', tint=0.5))
# 蓝色
Font(color='0000FF')