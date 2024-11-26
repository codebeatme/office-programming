# 创建只写工作簿，并添加工作表 Images1，Images2
from openpyxl import Workbook
wb = Workbook(True)
ws1 = wb.create_sheet('Images1')
ws2 = wb.create_sheet('Images2')

from openpyxl.drawing.image import Image

# 创建一个 Image 对象，并添加至 Images1
img = Image('python.png')
ws1.add_image(img, 'A1')
# 将 Image 对象添加至 Images2
ws2.add_image(img, 'B2')

wb.save('Diff.xlsx')

from openpyxl import load_workbook
# ERROR 不能以非只读方式读取 Diff.xlsx
wb = load_workbook('Diff.xlsx')
