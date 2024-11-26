# 讀取 Excel 檔案 Img.xlsx 中的工作表 Images
from openpyxl import load_workbook
wb = load_workbook('Img.xlsx')
ws = wb['Images']

# 取得 Excel 工作表中的影像的原始大小
for i in ws._images:
    print(f'原始大小：{i.width}x{i.height}')
