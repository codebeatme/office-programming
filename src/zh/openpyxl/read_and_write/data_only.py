# 请将命令行跳转至 Hello.xlsx 所在的目录，然后运行此脚本文件
# 导入函数 load_workbook
from openpyxl import load_workbook

# 使用 open 函数打开 Excel 文件
xlsx = open('Hello.xlsx', 'rb')
workbook = load_workbook(xlsx, data_only=True)

# 读取单元格 B4，C4 的公式计算结果并显示
worksheet = workbook['1.1班']
avg_age = worksheet['B4'].value
avg_score = worksheet['C4'].value
print(f'平均值 {avg_age} {avg_score}')
