# 打开 Link.xlsx 中的工作表 Hyperlinks
import openpyxl
workbook = openpyxl.load_workbook('Link.xlsx')
worksheet = workbook['Hyperlinks']

# 删除 A1 单元格中的链接
worksheet['A1'].hyperlink = None

workbook.save('Remove.xlsx')