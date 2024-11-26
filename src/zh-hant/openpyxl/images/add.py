# 建立唯寫活頁簿，並新增工作表 Images
from openpyxl import Workbook
wb = Workbook(True)
ws = wb.create_sheet('Images')

from openpyxl.drawing.image import Image

# 將影像 python.png 新增至工作表，錨點為 B2
img = Image('python.png')
ws.add_image(img, 'B2')

wb.save('Add.xlsx')