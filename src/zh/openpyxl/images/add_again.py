# 创建只写工作簿，并添加工作表 Images
from openpyxl import Workbook
wb = Workbook(True)
ws = wb.create_sheet('Images')

from openpyxl.drawing.image import Image

# 创建一个 Image 对象，并添加至 Images
img = Image('python.png')
ws.add_image(img, 'A1')
# 重复添加会产生警告，但不影响添加
ws.add_image(img, 'B2')
ws.add_image(img, 'C3')

wb.save('Again.xlsx')
