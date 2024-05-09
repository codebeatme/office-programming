# 请将命令行跳转至 Hello.xlsx 所在的目录，然后运行此脚本文件
# 导入包 openpyxl
import openpyxl

# 以只读方式打开 Excel 文件，并修改工作表标题
workbook = openpyxl.open('Hello.xlsx', True)
# 可以修改工作表标题
worksheet = workbook['1.1班']
worksheet.title = '1.2班'

# ERROR 只读的 Workbook 对象不能修改单元格的值
worksheet['A1'].value = '一个好人'
