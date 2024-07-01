# 打开 Data.xlsx 中的工作表 DateTime，并读取公式的计算结果
import openpyxl
workbook = openpyxl.load_workbook('Data.xlsx', read_only=True)
worksheet = workbook['DateTime']

# 用于显示单元格信息的函数
def show(cell):
    print(f'{cell.data_type} {cell.internal_value} {type(cell.internal_value)}')

show(worksheet['A1'])
show(worksheet['A2'])
show(worksheet['A3'])
