# 請將命令列跳躍至 Goods.xlsx 所在的目錄，然後執行此腳本檔案
from openpyxl import load_workbook

workbook = load_workbook('Goods.xlsx')
# 修改工作表 Tables 的名稱，並儲存為 Name.xlsx
workbook['Tables'].title = 'New Tables'
workbook.save('Name.xlsx')

workbook = load_workbook('Goods.xlsx')
# 隱藏工作表 Pens，Cups
workbook['Pens'].sheet_state = 'hidden'
workbook['Cups'].sheet_state = 'veryHidden'
workbook.save('Hidden.xlsx')

from openpyxl import Workbook

worksheet = Workbook(True).create_sheet()
worksheet.title = 'MySheet'
# 通過 parent 屬性取得工作表對應的活頁簿物件
worksheet.parent.save('Parent.xlsx')
