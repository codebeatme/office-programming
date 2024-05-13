# 請將命令列跳躍至 Hello.xlsx 所在的目錄，然後執行此腳本檔案
# 匯入函式 load_workbook
from openpyxl import load_workbook

# 使用 open 函式開啟 Excel 檔案
xlsx = open('Hello.xlsx', 'rb')
workbook = load_workbook(xlsx, data_only=True)

# 讀取儲存格 B4，C4 的公式計算結果並顯示
worksheet = workbook['1.1班']
avg_age = worksheet['B4'].value
avg_score = worksheet['C4'].value
print(f'平均值 {avg_age} {avg_score}')
