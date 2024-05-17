import openpyxl
workbook = openpyxl.Workbook()

# 由於已經存在 Sheet，因此新工作表名稱為 Sheet1，他將被放置在末尾
new_sheet1 = workbook.create_sheet()
print(f'新工作表名稱 {new_sheet1.title}')
print(f'最後一個工作表的名稱 {workbook.worksheets[-1].title}')

# 新工作表名稱為 Sheet2，他將被放置在開頭
workbook.create_sheet(index=0)
print(f'第一個工作表的名稱 {workbook.worksheets[0].title}')

# 新工作表名稱為 New，他將被放置在目前倒數第二個工作表 Sheet 之前
new = workbook.create_sheet('New', -2)
print(workbook.worksheets)
