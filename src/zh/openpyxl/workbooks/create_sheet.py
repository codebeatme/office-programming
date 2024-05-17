import openpyxl
workbook = openpyxl.Workbook()

# 由于已经存在 Sheet，因此新工作表名称为 Sheet1，他将被放置在末尾
new_sheet1 = workbook.create_sheet()
print(f'新工作表名称 {new_sheet1.title}')
print(f'最后一个工作表的名称 {workbook.worksheets[-1].title}')

# 新工作表名称为 Sheet2，他将被放置在开头
workbook.create_sheet(index=0)
print(f'第一个工作表的名称 {workbook.worksheets[0].title}')

# 新工作表名称为 New，他将被放置在目前倒数第二个工作表 Sheet 之前
new = workbook.create_sheet('New', -2)
print(workbook.worksheets)
