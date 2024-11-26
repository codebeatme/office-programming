# 读取 Excel 文件 Img.xlsx 中的工作表 Images
from openpyxl import load_workbook
wb = load_workbook('Img.xlsx')
ws = wb['Images']

for i in ws._images:
    # 获取 Excel 工作表中的图像的锚点
    a = i.anchor
    print(type(a))
    print(f'锚点：{a._from.col}，{a._from.row}')

    # 设置新的锚点
    i.anchor = 'E5'

wb.save('Anchor.xlsx')