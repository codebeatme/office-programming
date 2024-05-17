# 匯入 Workbook 類別
from openpyxl import Workbook

# 建立一個 Workbook 物件，將其中的工作表設定為隱藏
workbook = Workbook()
workbook['Sheet'].sheet_state = 'hidden'
workbook.create_sheet().sheet_state = 'hidden'
# ERROR 無法儲存沒有可見工作表的 Workbook 物件
workbook.save('nosheet.xlsx')
