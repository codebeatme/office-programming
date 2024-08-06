# 開啟 Data.xlsx 中的工作表 Students
from openpyxl import load_workbook
wb = load_workbook('Data.xlsx')
ws = wb['Students']

from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.cell_range import MultiCellRange, CellRange

# 為 C1:C1001 新增資料驗證，只能輸入在 0 之間 100 的整數
newDV = DataValidation(
    'whole', 0, 100,
    sqref='C1:C1001', imeMode='off', operator='between'
    )
ws.add_data_validation(newDV)

# 為 D1:D2，D4:D5 兩個範圍新增資料驗證，內容長度需要小於等於 5
ws.data_validations.append(DataValidation(
    'textLength', 5,
    sqref=(CellRange('D1:D2'), CellRange('D4:D5')), operator='lessThanOrEqual'
    ))

# 為 E1:E2，E4:E5 兩個範圍新增資料驗證，小數需要大於 1.5
ws.add_data_validation(DataValidation(
    'decimal', 1.5,
    sqref=MultiCellRange(['E1:E2', 'E4:E5']), operator='greaterThan'
    ))

wb.save('Add.xlsx')