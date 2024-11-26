# 建立唯寫活頁簿，並新增工作表 Images
from openpyxl import Workbook
wb = Workbook(True)
ws = wb.create_sheet('Images')

from openpyxl.drawing.image import Image

# 建立一個 Image 物件，並新增至 Images
img = Image('python.png')
ws.add_image(img, 'A1')
# 重複新增會產生警告，但不影響新增
ws.add_image(img, 'B2')
ws.add_image(img, 'C3')

wb.save('Again.xlsx')
