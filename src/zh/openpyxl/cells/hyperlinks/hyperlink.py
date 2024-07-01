# 打开 Data.xlsx 中的工作表 Hyperlinks
import openpyxl
workbook = openpyxl.load_workbook('Data.xlsx')
worksheet = workbook['Hyperlinks']

# 显示单元格 A1 的超链接的目标
link = worksheet['A1'].hyperlink
print(f'目标为：{link.target}')

# 删除单元格 B1 中的超链接
worksheet['B1'].hyperlink = None

# 单元格 C1 不包含超链接
print(worksheet['C1'].hyperlink)

# 修改超链接的目标
link.target = 'http://www.google.com/'
workbook.save('Google.xlsx')
