# 打开 Data.xlsx 中的工作表 Types，并读取公式的计算结果
import openpyxl
workbook = openpyxl.load_workbook('Data.xlsx', data_only=True)
worksheet = workbook['Types']

# 用于显示单元格信息的函数
def show(cell):
    print(f'{cell.data_type} {cell.value} {type(cell.value)}')

# 显示区域 A1:L1 中的单元格的信息
for row in worksheet['A1:L1']:
    for cell in row:
        show(cell)

# 将类型为 bytes 的值，传递给属性 value
worksheet['A2'].value = bytes(b'A good day')
show(worksheet['A2'])
