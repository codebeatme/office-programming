# 请将命令行跳转至 Goods.xlsx 所在的目录，然后运行此脚本文件
from openpyxl import load_workbook

workbook = load_workbook('Goods.xlsx')
worksheet = workbook.create_sheet()
# 以不同形式为工作表添加 4 行数据
worksheet.append((1.6, 2.5))
# 这相当于将单元格 A1 和 B1 移动至 A2 和 B2 的所在位置
worksheet.append([worksheet['A1'], worksheet['B1']])
worksheet.append({2: 1.0, 4: 2.0})
worksheet.append({'C': 1.1, 'E': 2.4})
# 保存为 Append.xlsx
workbook.save('Append.xlsx')

# 显示所有单元格的值
for r in worksheet.values:
    print(r)


from openpyxl import Workbook

workbook = Workbook(True)
worksheet = workbook.create_sheet()
# 为只写工作表添加数据，并保存为 AppendWriteOnly.xlsx
worksheet.append((1.6, 2.5))
worksheet.append([1.8, 2.2])
workbook.save('AppendWriteOnly.xlsx')
