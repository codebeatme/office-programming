# 匯入函式 save_workbook 和類別 Workbook
from openpyxl.writer.excel import save_workbook
from openpyxl import Workbook

# 建立 Workbook 物件並儲存
workbook = Workbook(True)
workbook.save('New.xlsx')
# ERROR 只能儲存一次
save_workbook(workbook, 'New.xlsx')