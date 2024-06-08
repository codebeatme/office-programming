# 请将命令行跳转至 Goods.xlsx 所在的目录，然后运行此脚本文件
from openpyxl import load_workbook

empty = load_workbook('Goods.xlsx')['Empty']
# C3 是一个值为空但拥有背景颜色的单元格
print(empty['C3'])
print(f'最小区域 {empty.dimensions}')
# 访问 E5 导致最小区域改变
empty['E5']
print(f'访问 E5 后的最小区域 {empty.calculate_dimension()}')

r_empty = load_workbook('Goods.xlsx', True)['Empty']
# 对于只读工作表，访问 E5 不会导致最小区域改变
r_empty['E5']
print(f'只读工作表访问 E5 后的最小区域 {r_empty.calculate_dimension()}')

# 重置最大行和最大列
r_empty.reset_dimensions()
# ERROR 需要将 force 参数设置为 True
r_empty.calculate_dimension()