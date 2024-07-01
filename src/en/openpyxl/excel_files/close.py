# Please jump the command line to the directory where Hello.xlsx is located and then run this script file
from openpyxl import Workbook, load_workbook

# Create a write-only workbook
w_workbook = Workbook(True)
w_workbook.create_sheet()
w_workbook['Sheet'].append(['Hello', 'World'])
# After calling the close method, write another row of data and save it
w_workbook.close()
w_workbook['Sheet'].append(['Hello', 'World'])
w_workbook.save('Close.xlsx')

# Create a read-only workbook
r_workbook = load_workbook('Hello.xlsx', True)
print(r_workbook['Class 1.1']['A1'].value)
# After the close method is called, cell A1 is read
r_workbook.close()
# ERROR Unable to read cell A1
print(r_workbook['Class 1.1']['A1'])