# 打开 Data.xlsx 中的工作表 Students
from openpyxl import load_workbook
workbook = load_workbook('Data.xlsx')
worksheet = workbook['Students']

dvs = worksheet.data_validations.dataValidation
# 显示 Excel 工作表中所有的数据验证
for dv in dvs:
    print(dv)

# 修改第二个数据验证
dv = dvs[1]
dv.errorTitle = '请选择一项'
dv.error = '只能选择男或女'
workbook.save('Save.xlsx')
