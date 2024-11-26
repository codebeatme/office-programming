# 讀取 Excel 檔案 Img.xlsx 中的工作表 Chat
from openpyxl import load_workbook
wb = load_workbook('Img.xlsx')
ws = wb['Chat']

# 取得 Excel 工作表中的影像的格式
for i in ws._images:
    print(f'格式：{i.format}')
