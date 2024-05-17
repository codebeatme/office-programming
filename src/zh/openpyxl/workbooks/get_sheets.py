# 请将命令行跳转至 School.xlsx 所在的目录，然后运行此脚本文件
import openpyxl
workbook = openpyxl.load_workbook('School.xlsx')

# 获取工作表 ClassA
sheet = workbook['ClassA']
print(sheet)
# 获取第二个和其之后的所有工作表
sheets = workbook.worksheets[1:]
print(sheets)
