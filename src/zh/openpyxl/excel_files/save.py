# 导入包 openpyxl
import openpyxl

# 创建 Workbook 并先后保存两次
wb = openpyxl.Workbook()
wb.save('overwrite.xlsx')
wb['Sheet']['A1'].value = '你好！'
wb.save('overwrite.xlsx')
