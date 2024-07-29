# 開啟 Link.xlsx 中的工作表 Hyperlinks
import openpyxl
wb = openpyxl.load_workbook('Link.xlsx')
ws = wb['Hyperlinks']

# 顯示範圍 A1:A2 中的儲存格的連結
for row in ws['A1:A2']:
    for cell in row:
        print(f'{cell.hyperlink}, {cell.hyperlink.target}')

from openpyxl.worksheet.hyperlink import Hyperlink
# B1 儲存格不會顯示為 Google，而是 https://www.google.com/，參數 ref 並不正確，但沒有關系
ws['B1'].hyperlink = Hyperlink('XXXX', display='Google', target='https://www.google.com/')
# 直接通過網址來設定 Excel 連結
ws['C1'].hyperlink = 'https://www.python.org'
# 指向工作表 Other 的 A1 儲存格的連結
ws['D1'].hyperlink = Hyperlink('D1', location='Other!A1')
# 指向本工作表的 A1 儲存格的連結
ws['E1'].hyperlink = Hyperlink('E1', location='A1')

wb.save('Google.xlsx')
