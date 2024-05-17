# 请将命令行跳转至 School.xlsx 所在的目录，然后运行此脚本文件
from openpyxl import load_workbook
workbook = load_workbook('School.xlsx')
print(f'当前活动 {workbook.active}')

# 将倒数第一个工作表设置为当前活动的工作表
workbook.active = -1
print(f'当前活动 {workbook.active}')
# 将 ClassB 设置为当前活动的工作表
workbook.active = workbook['ClassB']
print(f'当前活动 {workbook.active}')

# 没有索引为 100 的工作表
workbook.active = 100
print(f'当前活动 {workbook.active}')
