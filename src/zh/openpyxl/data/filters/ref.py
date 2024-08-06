# 打开 Data.xlsx 中的工作表 Trees
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Trees']

# 设置对 A，B 两列进行自动筛选
worksheet.auto_filter.ref = 'A1:B1'

workbook.save('Ref.xlsx')