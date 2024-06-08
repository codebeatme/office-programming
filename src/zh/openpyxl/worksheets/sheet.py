# 请将命令行跳转至 Goods.xlsx 所在的目录，然后运行此脚本文件
from openpyxl import load_workbook

workbook = load_workbook('Goods.xlsx')
# 修改工作表 Tables 的名称，并保存为 Name.xlsx
workbook['Tables'].title = 'New Tables'
workbook.save('Name.xlsx')

workbook = load_workbook('Goods.xlsx')
# 隐藏工作表 Pens，Cups
workbook['Pens'].sheet_state = 'hidden'
workbook['Cups'].sheet_state = 'veryHidden'
workbook.save('Hidden.xlsx')

from openpyxl import Workbook

worksheet = Workbook(True).create_sheet()
worksheet.title = 'MySheet'
# 通过 parent 属性获取工作表对应的工作簿对象
worksheet.parent.save('Parent.xlsx')

workbook = load_workbook('Goods.xlsx')
protection = workbook['Flowers'].protection
# 启用对工作表 Flowers 的保护
protection.sheet = True
# 允许用户在 Office 软件中删除行或列
protection.deleteColumns = False
protection.deleteRows = False
workbook.save('Protection.xlsx')

