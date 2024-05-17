# 请将命令行跳转至 School.xlsx 所在的目录，然后运行此脚本文件
from openpyxl import load_workbook
workbook = load_workbook('School.xlsx')

# 获取工作表的名称和索引
print(f'第二个工作表的名称 {workbook.sheetnames[1]}')
print(f'工作表 ClassC 的索引 {workbook.index(workbook["ClassC"])}')