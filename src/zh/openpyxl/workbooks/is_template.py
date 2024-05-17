# 请将命令行跳转至 Hello.xltx 所在的目录，然后运行此脚本文件
from openpyxl import load_workbook
workbook = load_workbook('Hello.xltx')

# 工作簿是否为模板？
print('template' in workbook.mime_type)