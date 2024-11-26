from openpyxl.styles import Color

# 红色，饱和度 50%
Color(rgb='FF0000', tint=0.5)
# 索引为 2 的索引颜色是红色
Color(indexed=2)
# 自动颜色
Color(auto=True)

from openpyxl.styles.colors import BLACK
# 黑色
Color(rgb=BLACK)
