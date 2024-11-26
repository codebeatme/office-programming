# 读取 Excel 文件 Img.xlsx 中的工作表 Images
from openpyxl import load_workbook
wb = load_workbook('Img.xlsx')
ws = wb['Images']

from openpyxl.drawing.image import Image

# 获取 Excel 工作表中的图像
print(f'一共 {len(ws._images)} 个图像')

# 替换第一个图像
ws._images[0] = Image('imac-icon.png')
# 删除第二个图像
del ws._images[1]
# 添加一个图像
ws._images.append(Image('python.png'))

wb.save('NewImg.xlsx')
