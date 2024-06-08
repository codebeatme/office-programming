# 请将命令行跳转至 Goods.xlsx 所在的目录，然后运行此脚本文件
from openpyxl import open

workbook = open('Goods.xlsx')
worksheet = workbook['Pens']
# 移动包含数据和公式单元格，并转换公式
worksheet.move_range('A1:C3', 1, 1, True)
workbook.save('Move.xlsx')