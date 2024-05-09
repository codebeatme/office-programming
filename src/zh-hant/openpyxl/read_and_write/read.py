# 請將命令列跳躍至 Hello.xlsx 所在的目錄，然後執行此腳本檔案
# 匯入函式 load_workbook
from openpyxl import load_workbook

# 讀取 Excel 檔案
workbook = load_workbook('Hello.xlsx')

# 取得工作表 1.1班 中的儲存格 A1，B1，C1，B4，C4 並顯示
worksheet = workbook['1.1班']
name = worksheet['A1'].value
age = worksheet['B1'].value
score = worksheet['C1'].value
print(f'第一個學生 {name} {age} {score}')
avg_age = worksheet['B4'].value
avg_score = worksheet['C4'].value
print(f'平均值公式 {avg_age} {avg_score}')
