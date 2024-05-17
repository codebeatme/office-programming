# 请将命令行跳转至 Hello.xlsx 所在的目录，然后运行此脚本文件
# 导入函数 open
from openpyxl import open, Workbook
wb = Workbook()
wb.get_sheet_by_name
# print(f'default {wb.active}')
# wb.active = 0
# print(f'0 {wb.active}')
# wb.active = 1
# print(f'1 {wb.active}')
# wb.active = -1
# print(f'-1 {wb.active}')
# wb.active = -2
# print(f'-2 {wb.active}')

wb2 = Workbook()
# ws2 = wb2.create_sheet()
# ERROR
# wb.active = ws2

# workbook = open('Hello.xlsx', True)
# worksheet = workbook.active
# print(workbook._active_sheet_index)

# print(f'wb["Sheet"] {wb["Sheet"]}')
# ERROR
# print(f'wb[1] {wb[0]}')

# print(f'active {wb.active.title}')
# wb.create_sheet('New Title 0', 0)
# print(f'active {wb.active.title}')
# print(f'0 {wb.sheetnames[0]}')
# print(f'1 {wb.sheetnames[1]}')
# print(f'active {wb.active.title}')

# wb.create_sheet('New Title 0', -2)
# print(f'active {wb.active.title}')
# print(f'0 {wb.sheetnames[0]}')
# print(f'1 {wb.sheetnames[1]}')
# print(f'2 {wb.sheetnames[2]}')
# print(f'-1 {wb.sheetnames[-1]}')
# print(f'-2 {wb.sheetnames[-2]}')
# print(f'-3 {wb.sheetnames[-3]}')
# print(f'active {wb.active.title}')

# print(f'new {wb.create_sheet('new 1').title}')
# print(f'new {wb.create_sheet('new 2').title}')
# print(f'-1 {wb.sheetnames[-1]}')
# print(f'-2 {wb.sheetnames[-2]}')

# print(f'add {wb.create_sheet('1 add at 0', 0).title}')
# print(f'add {wb.create_sheet('2 add at 0', 0).title}')
# print(f'0 {wb.sheetnames[0]}')
# print(f'1 {wb.sheetnames[1]}')

# print(f'add {wb.create_sheet('1 add at 100', 100).title}')
# print(f'add {wb.create_sheet('2 add at 100', 100).title}')
# print(f'-1 {wb.sheetnames[-1]}')
# print(f'-2 {wb.sheetnames[-2]}')

# print(f'add {wb.create_sheet('1 add at -100', -100).title}')
# print(f'add {wb.create_sheet('2 add at -100', -100).title}')
# print(f'0 {wb.sheetnames[0]}')
# print(f'1 {wb.sheetnames[1]}')

# ns = wb.create_sheet('New Sheet')
# cs = wb.copy_worksheet(ns)
# cs2 = wb.copy_worksheet(wb2['Sheet'])

# print(f'0 {wb.sheetnames[0]}')
# print(f'1 {wb.sheetnames[1]}')
# print(f'2 {wb.sheetnames[2]}')

# wb.create_sheet()
# ws = wb.create_sheet()
# wb.create_sheet()
# wb.create_sheet()
# wb.move_sheet(ws, -100)

# print(f'0 {wb.sheetnames[0]}')
# print(f'1 {wb.sheetnames[1]}')
# print(f'2 {wb.sheetnames[2]}')
# print(f'3 {wb.sheetnames[3]}')
# print(f'4 {wb.sheetnames[4]}')

# wb.remove(wb['Sheet'])
# del wb['Sheet']

# print(wb.path)

# wb.create_sheet()
# wb.create_sheet()
# wb.create_sheet()
# wb.create_sheet()
# print(wb.worksheets[5])
# print(wb.worksheets['Sheet'])

# print(wb.sheetnames[0])
# print(wb.index(wb['Sheet']))

from openpyxl import load_workbook
wb3 = load_workbook('Hello.xltx')
wb32 = load_workbook('Hello.xlsm')

# print(wb3.code_name)

# print(wb3.encoding)
# print(wb3.path)

# print(wb3.is_template)

wb4 = load_workbook('1904.xlsx')
# print(wb4.excel_base_date)
# print(wb4.epoch)
# print(wb3.mime_type)
# print(wb32.mime_type)
# print(wb4.mime_type)
