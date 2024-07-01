# 開啟 Data.xlsx 中的工作表 Types，並讀取公式的計算結果
import openpyxl
workbook = openpyxl.load_workbook('Data.xlsx', data_only=True)
worksheet = workbook['Types']

# 用於顯示儲存格資訊的函式
def show(cell):
    print(f'{cell.data_type} {cell.value} {type(cell.value)}')

# 顯示區域 A1:L1 中的儲存格的資訊
for row in worksheet['A1:L1']:
    for cell in row:
        show(cell)

# 將型別為 bytes 的值，傳遞給屬性 value
worksheet['A2'].value = bytes(b'A good day')
show(worksheet['A2'])
