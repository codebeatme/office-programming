# 创建只写工作簿
from openpyxl import Workbook
wb = Workbook(True)

from openpyxl.styles import Font
from openpyxl.styles.named_styles import NamedStyle

# 创建一个命名格式
ns = NamedStyle(
    '绿色主题',
    # 字体，绿色
    font=Font(color='00FF00'),
    # 数字显示为百分比
    number_format='0%'
)

# 将命名格式添加至 Excel 工作簿
wb.add_named_style(ns)

wb.save('Add.xlsx')