# Import function save_workbook and class Workbook
from openpyxl.writer.excel import save_workbook
from openpyxl import Workbook

# Create the Workbook object and save it
workbook = Workbook(True)
workbook.save('New.xlsx')
# ERROR Can only be saved once
save_workbook(workbook, 'New.xlsx')