# Please jump the command line to the directory where Hello.xlsx is located and then run this script file
# Import package openpyxl
import openpyxl

# Open an Excel file in read-only mode and modify the worksheet title
workbook = openpyxl.open('Hello.xlsx', True)
worksheet = workbook['Class 1.1']
worksheet.title = 'Class 1.2'

# ERROR Read-only Workbook objects can't modify cell values.
worksheet['A1'].value = 'A good man'