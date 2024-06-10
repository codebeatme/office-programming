# 读取 Excel 文件 Food.xlsx 中的工作表 Bread
import openpyxl
workbook = openpyxl.load_workbook('Food.xlsx')
worksheet = workbook['Bread']

# 在当前第一行的位置插入两行
worksheet.insert_rows(1, 2)
# 在当前第一列的位置插入两列
worksheet.insert_cols(1, 2)
# 删除第四行和第四列
worksheet.delete_rows(4)
worksheet.delete_cols(4)

# 保存为 Excel 文件 NewFood.xlsx
workbook.save('NewFood.xlsx')