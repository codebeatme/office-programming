# 请将命令行跳转至 Goods.xlsx 所在的目录，然后运行此脚本文件
import openpyxl

workbook = openpyxl.load_workbook('Goods.xlsx', True)
phones = workbook['Phones']
# 获取单元格 A1，C1
print(phones.cell(1, 1))
print(phones['C1'])
# 获取第二和第四行之间，从第一列开始至最大列结束的区域内的单元格
print(phones[2:4])

cups = workbook['Cups']
# 工作表末尾的一些行将被忽略
print(cups['A1:B10'])
