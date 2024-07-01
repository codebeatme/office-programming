# 開啟 Link.xlsx 中的工作表 Hyperlinks
import openpyxl
workbook = openpyxl.load_workbook('Link.xlsx')
worksheet = workbook['Hyperlinks']

# 刪除 A1 儲存格中的連結
worksheet['A1'].hyperlink = None

workbook.save('Remove.xlsx')