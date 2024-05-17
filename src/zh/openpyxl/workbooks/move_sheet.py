# 请将命令行跳转至 School.xlsx 所在的目录，然后运行此脚本文件
from openpyxl import open
workbook = open('School.xlsx')
print(workbook.worksheets)

# 工作表 ClassA 将位于末尾
workbook.move_sheet('ClassA', 100)
print(workbook.worksheets)
# 当前第一个工作表 ClassB 将被移动至倒数第二的位置
b = workbook['ClassB']
workbook.move_sheet(b, -1)
print(workbook.worksheets)
