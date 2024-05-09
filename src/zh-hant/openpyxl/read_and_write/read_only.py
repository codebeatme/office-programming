# 請將命令列跳躍至 Hello.xlsx 所在的目錄，然後執行此腳本檔案
# 匯入套件 openpyxl
import openpyxl

# 以唯讀方式開啟 Excel 檔案，並修改工作表標題
workbook = openpyxl.open('Hello.xlsx', True)
# 可以修改工作表標題
worksheet = workbook['1.1班']
worksheet.title = '1.2班'

# ERROR 唯讀的 Workbook 物件不能修改儲存格的值
worksheet['A1'].value = '一個好人'
