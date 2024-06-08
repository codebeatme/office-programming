# Overwrite.xlsx 會被儲存至命令列的目前工作目錄
# 匯入套件 openpyxl
import openpyxl

# 建立 Workbook 並先後儲存兩次
workbook = openpyxl.Workbook()
workbook['Sheet']['A1'].value = '世界'
workbook.save('Overwrite.xlsx')
workbook['Sheet']['A1'].value = '你好！'
workbook.save('Overwrite.xlsx')
