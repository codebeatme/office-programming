from openpyxl import Workbook

# 创建一个新的工作簿，默认包含 Sheet 工作表
workbook = Workbook()
print(workbook['Sheet'])

# 只写工作簿不包含任何工作表
write_only_workbook = Workbook(True)
print(write_only_workbook.worksheets)