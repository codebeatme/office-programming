# 创建只写工作簿，并添加工作表 Images
from openpyxl import Workbook
wb = Workbook(True)
ws = wb.create_sheet('Images')

from openpyxl.drawing.image import Image

# 将图像 python.png 添加至工作表，锚点为 B2
img = Image('python.png')
ws.add_image(img, 'B2')

wb.save('Add.xlsx')