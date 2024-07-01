# 開啟 Link.xlsx 中的工作表 Hyperlinks
import openpyxl
workbook = openpyxl.load_workbook('Link.xlsx')
worksheet = workbook['Hyperlinks']

link = worksheet['A2'].hyperlink
# 修改工具提示和指向的儲存格位址
link.tooltip = '一個大概不會顯示的工具提示'
link.location = 'Other!A1'
# 修改 display 不影響儲存格的內容，因此修改 value 屬性
worksheet['A2'].value = '指向 Other 的 A1'
link.display = '不影響儲存格內容'

workbook.save('Modify.xlsx')