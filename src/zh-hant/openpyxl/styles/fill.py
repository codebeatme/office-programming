from openpyxl.styles import PatternFill, GradientFill, Color
from openpyxl.styles.fills import Stop

# 圖樣 darkDown，前景紅色，背景綠色
PatternFill('darkDown', 'FF0000', Color('00FF00'))
# 旋轉 90 度，開始和結束為紅色，中間為綠色
GradientFill('path', 90, stop=(
    Stop('FF0000', 0),
    Stop('00FF00', 0.5),
    Stop('FF0000', 1),
))