from openpyxl import load_workbook

wb = load_workbook('Hello.xlsx')
ws = wb['Sheet']
ws['C3']

for r in wb['Sheet'].rows:
    for c in r:
        print(c.value)