# Overwrite.xlsx will be saved to the current working directory of the command line
# Import package openpyxl
import openpyxl

# Create a workbook and save it twice
workbook = openpyxl.Workbook()
workbook.save('Overwrite.xlsx')
workbook['Sheet']['A1'].value = 'Hello!'
workbook.save('Overwrite.xlsx')