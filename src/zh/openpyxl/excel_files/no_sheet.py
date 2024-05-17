# 导入 Workbook 类
from openpyxl import Workbook

# 创建一个 Workbook 对象，将其中的工作表设置为隐藏
workbook = Workbook()
workbook['Sheet'].sheet_state = 'hidden'
workbook.create_sheet().sheet_state = 'hidden'
# ERROR 无法保存没有可见工作表的 Workbook 对象
workbook.save('nosheet.xlsx')
