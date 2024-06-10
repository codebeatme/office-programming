# 讀取 Excel 檔案 Food.xlsx 中的工作表 Fish
import openpyxl
workbook = openpyxl.load_workbook('Food.xlsx')
worksheet = workbook['Fish']

# 在工作表的開始插入兩列，原儲存格 B4 的計算公式不會改變
worksheet.insert_rows(1, 2)

# 儲存為 Excel 檔案 Formula.xlsx
workbook.save('Formula.xlsx')