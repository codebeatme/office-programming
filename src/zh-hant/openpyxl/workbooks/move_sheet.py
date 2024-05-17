# 請將命令列跳躍至 School.xlsx 所在的目錄，然後執行此腳本檔案
from openpyxl import open
workbook = open('School.xlsx')
print(workbook.worksheets)

# 工作表 ClassA 將位於末尾
workbook.move_sheet('ClassA', 100)
print(workbook.worksheets)
# 目前第一個工作表 ClassB 將被移動至倒數第二的位置
b = workbook['ClassB']
workbook.move_sheet(b, -1)
print(workbook.worksheets)
