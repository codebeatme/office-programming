# 請將命令列跳躍至 Goods.xlsx 所在的目錄，然後執行此腳本檔案
import openpyxl

workbook = openpyxl.load_workbook('Goods.xlsx', True)
phones = workbook['Phones']
# 取得儲存格 A1，C1
print(phones.cell(1, 1))
print(phones['C1'])
# 取得第二和第四列之間，從第一欄開始至最大欄結束的區域內的儲存格
print(phones[2:4])

cups = workbook['Cups']
# 工作表末尾的一些列將被忽略
print(cups['A1:B10'])
