# 請將命令列跳躍至 School.xlsx 所在的目錄，然後執行此腳本檔案
from openpyxl import open
workbook = open('School.xlsx')

# 刪除工作表 ClassA，ClassB
del workbook['ClassA']
workbook.remove(workbook['ClassB'])
print(workbook.worksheets)
