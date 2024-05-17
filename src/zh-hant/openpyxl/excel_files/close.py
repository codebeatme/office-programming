# 請將命令列跳躍至 Hello.xlsx 所在的目錄，然後執行此腳本檔案
from openpyxl import Workbook, load_workbook

# 建立唯寫的活頁簿
w_workbook = Workbook(True)
w_workbook.create_sheet()
w_workbook['Sheet'].append(['Hello', 'World'])
# 呼叫 close 方法之後，再次寫入一行資料，並儲存
w_workbook.close()
w_workbook['Sheet'].append(['你好', '世界'])
w_workbook.save('w_close.xlsx')

# 建立唯讀的活頁簿
r_workbook = load_workbook('Hello.xlsx', True)
print(r_workbook['1.1班']['A1'].value)
# 呼叫 close 方法之後，讀取儲存格 A1
r_workbook.close()
# ERROR 無法讀取 A1 儲存格
print(r_workbook['1.1班']['A1'])
