# 打开 Data.xlsx 中的工作表 Values
import openpyxl
workbook = openpyxl.load_workbook('Data.xlsx')
worksheet = workbook['Values']

# 显示一些单元格，A2 将成为一个空单元格
worksheet['A2'].value = None
print(f"A1：{type(worksheet['A1'].value)} {worksheet['A1'].value}")
print(f"B1：{type(worksheet['B1'].value)} {worksheet['B1'].value}")
print(f"C1：{type(worksheet['C1'].value)} {worksheet['C1'].value}")
print(f"D4：{type(worksheet['D4'].internal_value)} {worksheet['D4'].internal_value}")
workbook.save('Value.xlsx')

# 以只读方式打开 Data.xlsx 中的工作表 Values
r_workbook = openpyxl.load_workbook('Data.xlsx', True)
r_worksheet = r_workbook['Values']

# 读取单元格 A1，D4
print(f"A1：{r_worksheet['A1']} {type(r_worksheet['A1'].internal_value)} {r_worksheet['A1'].internal_value}")
print(f"D4：{r_worksheet['D4']} {type(r_worksheet['D4'].value)} {r_worksheet['D4'].value}")
