from openpyxl.styles import PatternFill, GradientFill, Color
from openpyxl.styles.fills import Stop

# 图案 darkDown，前景红色，背景绿色
PatternFill('darkDown', 'FF0000', Color('00FF00'))
# 旋转 90 度，开始和结束为红色，中间为绿色
GradientFill('path', 90, stop=(
    Stop('FF0000', 0),
    Stop('00FF00', 0.5),
    Stop('FF0000', 1),
))