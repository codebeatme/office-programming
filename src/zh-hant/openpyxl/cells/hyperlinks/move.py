# 開啟 Link.xlsx 中的工作表 Hyperlinks
import openpyxl
workbook = openpyxl.load_workbook('Link.xlsx')
worksheet = workbook['Hyperlinks']

# 移動儲存格 A1，A2 至 D1，D2
worksheet.move_range('A1:A2', cols=3)
# 修改 D1 儲存格的連結的 ref
worksheet['D1'].hyperlink.ref = 'D1'

workbook.save('Move.xlsx')