# 打开 Link.xlsx 中的工作表 Hyperlinks
import openpyxl
workbook = openpyxl.load_workbook('Link.xlsx')
worksheet = workbook['Hyperlinks']

link = worksheet['A2'].hyperlink
# 修改工具提示和指向的单元格地址
link.tooltip = '一个大概不会显示的工具提示'
link.location = 'Other!A1'
# 修改 display 不影响单元格的内容，因此修改 value 属性
worksheet['A2'].value = '指向 Other 的 A1'
link.display = '不影响单元格内容'

workbook.save('Modify.xlsx')