from openpyxl import Workbook

# 建立一個新的活頁簿，預設包含 Sheet 工作表
workbook = Workbook()
print(workbook['Sheet'])

# 唯寫活頁簿不包含任何工作表
write_only_workbook = Workbook(True)
print(write_only_workbook.worksheets)