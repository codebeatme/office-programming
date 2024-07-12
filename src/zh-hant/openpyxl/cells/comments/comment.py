# 開啟 Comment.xlsx 中的工作表 Sheet
import openpyxl
wb = openpyxl.load_workbook('Comment.xlsx')
ws = wb['Sheet']

from openpyxl.comments.comments import Comment
# 為儲存格 A1 設定註解
ws['A1'].comment = Comment('一個註解！', '好人', 300, 500)
# 取得儲存格 B1，C1 的註解
print(ws['B1'].comment)
print(ws['C1'].comment)
# 修改儲存格 C1 的註解
c = ws['C1'].comment
c.text = '我改了一下'
c.width = 500
c.height = 500

wb.save('Apple.xlsx')