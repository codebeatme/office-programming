# Import Workbook class
from openpyxl import Workbook

# Create a Workbook object and set the worksheets in it to be hidden
workbook = Workbook()
workbook['Sheet'].sheet_state = 'hidden'
workbook.create_sheet().sheet_state = 'hidden'
# ERROR Unable to save Workbook objects without visible worksheets
workbook.save('nosheet.xlsx')