# 读取 Excel 文件 Img.xlsx 中的工作表 Chat
from openpyxl import load_workbook
wb = load_workbook('Img.xlsx')
ws = wb['Chat']

# 获取 Excel 工作表中的图像的格式
for i in ws._images:
    print(f'格式：{i.format}')
