# 导入函数 save_workbook 和类 Workbook
from openpyxl.writer.excel import save_workbook
from openpyxl import Workbook

# 创建 Workbook 对象并保存
workbook = Workbook(True)
workbook.save('New.xlsx')
# ERROR 只能保存一次
save_workbook(workbook, 'New.xlsx')