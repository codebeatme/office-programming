# 讀取 Excel 檔案 Img.xlsx 中的工作表 Images
from openpyxl import load_workbook
wb = load_workbook('Img.xlsx')
ws = wb['Images']

from openpyxl.drawing.image import Image

# 取得 Excel 工作表中的影像
print(f'一共 {len(ws._images)} 個影像')

# 取代第一個影像
ws._images[0] = Image('imac-icon.png')
# 刪除第二個影像
del ws._images[1]
# 新增一個影像
ws._images.append(Image('python.png'))

wb.save('NewImg.xlsx')
