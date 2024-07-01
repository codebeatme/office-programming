# Please jump the command line to the directory where Hello.xlsx is located and then run this script file
# Import function load_workbook
from openpyxl import load_workbook

# Use the open function to open the Excel file
xlsx = open('Hello.xlsx', 'rb')
workbook = load_workbook(xlsx, data_only=True)

# Read and display the results of the formula in cells B4 and C4
worksheet = workbook['Class 1.1']
avg_age = worksheet['B4'].value
avg_score = worksheet['C4'].value
print(f'Average {avg_age} {avg_score}')