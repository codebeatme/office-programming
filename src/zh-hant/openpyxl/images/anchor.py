# 讀取 Excel 檔案 Img.xlsx 中的工作表 Images
from openpyxl import load_workbook
wb = load_workbook('Img.xlsx')
ws = wb['Images']

for i in ws._images:
    # 取得 Excel 工作表中的影像的錨點
    a = i.anchor
    print(type(a))
    print(f'錨點：{a._from.col}，{a._from.row}')

    # 設定新的錨點
    i.anchor = 'E5'

wb.save('Anchor.xlsx')