# 建立唯寫活頁簿，並新增工作表 Images1，Images2
from openpyxl import Workbook
wb = Workbook(True)
ws1 = wb.create_sheet('Images1')
ws2 = wb.create_sheet('Images2')

from openpyxl.drawing.image import Image

# 建立一個 Image 物件，並新增至 Images1
img = Image('python.png')
ws1.add_image(img, 'A1')
# 將 Image 物件新增至 Images2
ws2.add_image(img, 'B2')

wb.save('Diff.xlsx')

from openpyxl import load_workbook
# ERROR 不能以非唯讀方式讀取 Diff.xlsx
wb = load_workbook('Diff.xlsx')
