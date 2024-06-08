# 请将命令行跳转至 Goods.xlsx 所在的目录，然后运行此脚本文件
import openpyxl

worksheet = openpyxl.load_workbook('Goods.xlsx')['Phones']
# 访问单元格 A4 将导致工作表的最大行发生变化
worksheet['A4']
# 显示 Phones 工作表中的单元格的值
for row in worksheet.values:
    print(row)
