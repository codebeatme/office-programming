# 请将命令行跳转至 Goods.xlsx 所在的目录，然后运行此脚本文件
from openpyxl import open

workbook = open('Goods.xlsx')
worksheet = workbook['Flowers']
print(f'合并之前的 B2 为 {worksheet["B2"].value}')
# 合并区域 A1:B2 之后，立即取消合并
worksheet.merge_cells('A1:B2')
print(f'合并之后的 B2 为 {worksheet["B2"].value}')
worksheet.unmerge_cells(start_row=1, start_column=1, end_row=2, end_column=2)

# 合并区域 B2:C3，D4:H6
worksheet.merge_cells('B2:C3')
worksheet.merge_cells('D4:H6')
# 判断单元格 E3，D4 是否被合并了
print(f'单元格 E3 被合并了吗？{"E3" in worksheet.merged_cells}')
print(f'单元格 D4 被合并了吗？{"D4" in worksheet.merged_cells}')