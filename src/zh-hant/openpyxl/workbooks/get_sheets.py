# 請將命令列跳躍至 School.xlsx 所在的目錄，然後執行此腳本檔案
import openpyxl
workbook = openpyxl.load_workbook('School.xlsx')

# 取得工作表 ClassA
sheet = workbook['ClassA']
print(sheet)
# 取得第二個和其之後的所有工作表
sheets = workbook.worksheets[1:]
print(sheets)
