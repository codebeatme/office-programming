# Please jump the command line to the directory where Hello.xlsx is located and then run this script file
# Import function load_workbook
from openpyxl import load_workbook

# Read the Excel file
workbook = load_workbook('Hello.xlsx')

# Get cells A1, B1, C1, B4, C4 in Worksheet 1.1 and display them.
worksheet = workbook['Class 1.1']
name = worksheet['A1'].value
age = worksheet['B1'].value
score = worksheet['C1'].value
print(f'The first student {name} {age} {score}')
avg_age = worksheet['B4'].value
avg_score = worksheet['C4'].value
print(f'Average formula {avg_age} {avg_score}')