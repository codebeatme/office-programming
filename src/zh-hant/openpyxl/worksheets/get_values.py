# 請將命令列跳躍至 Goods.xlsx 所在的目錄，然後執行此腳本檔案
import openpyxl

worksheet = openpyxl.load_workbook('Goods.xlsx')['Phones']
# 存取儲存格 A4 將導致工作表的最大列發生變化
worksheet['A4']
# 顯示 Phones 工作表中的儲存格的值
for row in worksheet.values:
    print(row)
