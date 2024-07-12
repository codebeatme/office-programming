# 打开 Comment.xlsx 中的工作表 Sheet
import openpyxl
wb = openpyxl.load_workbook('Comment.xlsx')
ws = wb['Sheet']

from openpyxl.comments.comments import Comment
# 为单元格 A1 设置批注
ws['A1'].comment = Comment('一个批注！', '好人', 300, 500)
# 获取单元格 B1，C1 的批注
print(ws['B1'].comment)
print(ws['C1'].comment)
# 修改单元格 C1 的批注
c = ws['C1'].comment
c.text = '我改了一下'
c.width = 500
c.height = 500

wb.save('Apple.xlsx')