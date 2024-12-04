# 建立唯寫活頁簿
from openpyxl import Workbook
wb = Workbook(True)

from openpyxl.styles import Font
from openpyxl.styles.named_styles import NamedStyle

# 建立一個命名格式
ns = NamedStyle(
    '綠色主題',
    # 字型，綠色
    font=Font(color='00FF00'),
    # 數值顯示為百分比
    number_format='0%'
)

# 將命名格式新增至 Excel 活頁簿
wb.add_named_style(ns)

wb.save('Add.xlsx')