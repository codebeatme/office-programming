# 请将命令行跳转至 Goods.xlsx 所在的目录，然后运行此脚本文件
from openpyxl import load_workbook

workbook = load_workbook('Goods.xlsx')
protection = workbook['Flowers'].protection
# 启用对工作表 Flowers 的保护
protection.enable()
# 允许用户在 Office 软件中删除行或列
protection.deleteColumns = False
protection.deleteRows = False
# 设置密码
protection.set_password('123')

workbook.save('Protection.xlsx')
