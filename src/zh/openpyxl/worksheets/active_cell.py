# 请将命令行跳转至 Goods.xlsx 所在的目录，然后运行此脚本文件
from openpyxl import load_workbook

worksheet = load_workbook('Goods.xlsx')['Tables']
# 显示工作表当前活动的单元格的地址
print(f'当前活动的单元格为 {worksheet.active_cell}')
