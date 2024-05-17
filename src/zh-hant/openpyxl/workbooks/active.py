# 請將命令列跳躍至 School.xlsx 所在的目錄，然後執行此腳本檔案
from openpyxl import load_workbook
workbook = load_workbook('School.xlsx')
print(f'目前使用中 {workbook.active}')

# 將倒數第一個工作表設定為目前使用中的工作表
workbook.active = -1
print(f'目前使用中 {workbook.active}')
# 將 ClassB 設定為目前使用中的工作表
workbook.active = workbook['ClassB']
print(f'目前使用中 {workbook.active}')

# 沒有索引為 100 的工作表
workbook.active = 100
print(f'目前使用中 {workbook.active}')
