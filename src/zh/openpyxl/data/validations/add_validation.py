# 打开 Data.xlsx 中的工作表 Students
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Students']

from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.cell_range import MultiCellRange, CellRange

# 为 C1:C1001 添加数据验证，只能输入在 0 之间 100 的整数
newDV = DataValidation(
    'whole', 0, 100,
    sqref='C1:C1001', imeMode='off', operator='between'
    )
ws.add_data_validation(newDV)

# 为 D1:D2，D4:D5 两个区域添加数据验证，内容长度需要小于等于 5
ws.data_validations.append(DataValidation(
    'textLength', 5,
    sqref=(CellRange('D1:D2'), CellRange('D4:D5')), operator='lessThanOrEqual'
    ))

# 为 E1:E2，E4:E5 两个区域添加数据验证，小数需要大于 1.5
ws.add_data_validation(DataValidation(
    'decimal', 1.5,
    sqref=MultiCellRange(['E1:E2', 'E4:E5']), operator='greaterThan'
    ))

wb.save('Add.xlsx')