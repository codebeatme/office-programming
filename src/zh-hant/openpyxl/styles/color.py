from openpyxl.styles import Color

# 紅色，飽和度 50%
Color(rgb='FF0000', tint=0.5)
# 索引為 2 的索引色彩是紅色
Color(indexed=2)
# 自動色彩
Color(auto=True)

from openpyxl.styles.colors import BLACK
# 黑色
Color(rgb=BLACK)
