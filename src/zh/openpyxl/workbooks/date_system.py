# 请将命令行跳转至 1904.xlsx 所在的目录，然后运行此脚本文件
from openpyxl import load_workbook
workbook = load_workbook('1904.xlsx')

# 获取工作簿的日期系统
print(workbook.excel_base_date)