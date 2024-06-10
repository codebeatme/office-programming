# 读取 Excel 文件 Food.xlsx 中的工作表 Fish
import openpyxl
workbook = openpyxl.load_workbook('Food.xlsx')
worksheet = workbook['Fish']

# 在工作表的开始插入两行，原单元格 B4 的计算公式不会改变
worksheet.insert_rows(1, 2)

# 保存为 Excel 文件 Formula.xlsx
workbook.save('Formula.xlsx')