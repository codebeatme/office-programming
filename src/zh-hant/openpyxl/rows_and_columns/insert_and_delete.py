# 讀取 Excel 檔案 Food.xlsx 中的工作表 Bread
import openpyxl
workbook = openpyxl.load_workbook('Food.xlsx')
worksheet = workbook['Bread']

# 在目前第一列的位置插入兩列
worksheet.insert_rows(1, 2)
# 在目前第一欄的位置插入兩欄
worksheet.insert_cols(1, 2)
# 刪除第四列和第四欄
worksheet.delete_rows(4)
worksheet.delete_cols(4)

# 儲存為 Excel 檔案 NewFood.xlsx
workbook.save('NewFood.xlsx')