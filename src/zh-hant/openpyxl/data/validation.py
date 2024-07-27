# 開啟 Data.xlsx 中的工作表 Students
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Students']

dvs = worksheet.data_validations.dataValidation
# 顯示 Excel 工作表中所有的資料驗證
for dv in dvs:
    print(dv)

# 修改第二個資料驗證
dv = dvs[1]
dv.errorTitle = '請選擇一項'
dv.error = '只能選擇男或女'
workbook.save('Save.xlsx')
