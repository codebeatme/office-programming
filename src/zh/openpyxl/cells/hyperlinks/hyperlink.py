# 打开 Link.xlsx 中的工作表 Hyperlinks
import openpyxl
wb = openpyxl.load_workbook('Link.xlsx')
ws = wb['Hyperlinks']

# 显示区域 A1:A2 中的单元格的链接
for row in ws['A1:A2']:
    for cell in row:
        print(f'{cell.hyperlink}, {cell.hyperlink.target}')

from openpyxl.worksheet.hyperlink import Hyperlink
# B1 单元格不会显示为 Google，而是 https://www.google.com/，参数 ref 并不正确，但没有关系
ws['B1'].hyperlink = Hyperlink('XXXX', display='Google', target='https://www.google.com/')
# 直接通过网址来设置 Excel 链接
ws['C1'].hyperlink = 'https://www.python.org'
# 指向工作表 Other 的 A1 单元格的链接
ws['D1'].hyperlink = Hyperlink('D1', location='Other!A1')
# 指向本工作表的 A1 单元格的链接
ws['E1'].hyperlink = Hyperlink('E1', location='A1')

wb.save('Google.xlsx')
