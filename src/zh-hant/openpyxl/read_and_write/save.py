# 匯入套件 openpyxl
import openpyxl

# 建立 Workbook 並先後儲存兩次
wb = openpyxl.Workbook()
wb.save('overwrite.xlsx')
wb['Sheet']['A1'].value = '你好！'
wb.save('overwrite.xlsx')
