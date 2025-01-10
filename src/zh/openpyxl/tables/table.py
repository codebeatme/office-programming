# 读取 Excel 文件 Tables.xlsx 中的工作表 Tree
from openpyxl import load_workbook
wb = load_workbook(r'E:\cxc\Documents\code\beat\office-programming\src\zh\openpyxl\tables\Tables.xlsx')
ws = wb['Tree']

from openpyxl.worksheet.table import Table

ws.add_table(Table(displayName='ABC', ref='A1:C5', tableType='worksheet', headerRowCount=2))

# for t in ws.tables:
#     t1:Table = ws.tables.get(t)
#     print(t1)

wb.save('T.xlsx')

wb = load_workbook(r'T.xlsx')
ws = wb['Tree']

for t in ws.tables:
    t1:Table = ws.tables.get(t)
    print(t1)
