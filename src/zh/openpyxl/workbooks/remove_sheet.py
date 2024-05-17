# 请将命令行跳转至 School.xlsx 所在的目录，然后运行此脚本文件
from openpyxl import open
workbook = open('School.xlsx')

# 删除工作表 ClassA，ClassB
del workbook['ClassA']
workbook.remove(workbook['ClassB'])
print(workbook.worksheets)
