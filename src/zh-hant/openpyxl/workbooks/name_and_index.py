# 請將命令列跳躍至 School.xlsx 所在的目錄，然後執行此腳本檔案
from openpyxl import load_workbook
workbook = load_workbook('School.xlsx')

# 取得工作表的名稱和索引
print(f'第二個工作表的名稱 {workbook.sheetnames[1]}')
print(f'工作表 ClassC 的索引 {workbook.index(workbook["ClassC"])}')