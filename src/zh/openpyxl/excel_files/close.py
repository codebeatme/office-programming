# 请将命令行跳转至 Hello.xlsx 所在的目录，然后运行此脚本文件
from openpyxl import Workbook, load_workbook

# 创建只写的工作簿
w_workbook = Workbook(True)
w_workbook.create_sheet()
w_workbook['Sheet'].append(['Hello', 'World'])
# 调用 close 方法之后，再次写入一行数据，并保存
w_workbook.close()
w_workbook['Sheet'].append(['你好', '世界'])
w_workbook.save('w_close.xlsx')

# 创建只读的工作簿
r_workbook = load_workbook('Hello.xlsx', True)
print(r_workbook['1.1班']['A1'].value)
# 调用 close 方法之后，读取单元格 A1
r_workbook.close()
# ERROR 无法读取 A1 单元格
print(r_workbook['1.1班']['A1'])
