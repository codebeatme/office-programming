from openpyxl.styles import Color

# 红色，饱和度 50%
c = Color(rgb='FF0000', tint=0.5)
print(c.value)
# 将颜色设置为蓝色
c.value = '0000FF'

# 索引为 2 的索引颜色是红色
Color(indexed=2)
# 自动颜色
Color(auto=True)

from openpyxl.styles.colors import BLACK
# 黑色
Color(rgb=BLACK)
