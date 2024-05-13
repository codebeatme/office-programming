# 请将命令行跳转至 Hello.xlsx 所在的目录，然后运行此脚本文件
# 导入函数 load_workbook
from openpyxl import load_workbook

# 读取 Excel 文件
workbook = load_workbook('Hello.xlsx')

# 获取工作表 1.1班 中的单元格 A1，B1，C1，B4，C4 并显示
worksheet = workbook['1.1班']
name = worksheet['A1'].value
age = worksheet['B1'].value
score = worksheet['C1'].value
print(f'第一个学生 {name} {age} {score}')
avg_age = worksheet['B4'].value
avg_score = worksheet['C4'].value
print(f'平均值公式 {avg_age} {avg_score}')
