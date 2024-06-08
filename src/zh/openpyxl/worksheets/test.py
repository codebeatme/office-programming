from openpyxl import load_workbook, Workbook
import openpyxl
import openpyxl.worksheet
import openpyxl.worksheet._write_only
import openpyxl.worksheet._read_only
import openpyxl.worksheet.worksheet
from openpyxl.cell.cell import Cell

workbook = load_workbook('E:/cxc/Documents/code/beat/office-programming/src/zh/openpyxl/worksheets/Goods.xlsx')
worksheet:openpyxl.worksheet.worksheet.Worksheet = workbook['Tables']
workbookrr = load_workbook('E:/cxc/Documents/code/beat/office-programming/src/zh/openpyxl/worksheets/Goods.xlsx', True)
worksheetrr:openpyxl.worksheet._read_only.ReadOnlyWorksheet = workbookrr['Tables']
# worksheetrr.values
# for r in worksheetrr.values:
#     for c in r:
#         print(c)
# workbookrr.title = 'ABC'
# print(workbookrr.title)
# print(worksheetrr.calculate_dimension(True))
# for c in worksheetrr.rows:
#     print(c)

# worksheet.append([worksheet['A1'].value,worksheet['B1'].value])
# workbook.save('Append.xlsx')

w_workbook = Workbook(True)
ws:openpyxl.worksheet._write_only.WriteOnlyWorksheet = w_workbook.create_sheet()

# ws.append([1])
# ws.title = 'ABC'
# print(ws.title)
# print(ws.encoding)
# print(ws.path)

# print(worksheet['B1:F4'])
# print(worksheet.cell(1 , 1))
# print(worksheet['A1'])
# print(worksheet.cell(1 , 2, 1))
# print(worksheet['B1'])
# print(worksheet.cell(1 , 1) == worksheet['A1'])

# print(worksheet.active_cell)
# worksheet.active_cell = 'A1'

# worksheet.move_range('B1:D3', 3, translate=True)
# print(worksheet['D1'].value)
# print(worksheet['D3'].value)
# print(worksheet['D4'].internal_value)
# print(worksheet['D6'].internal_value)
# workbook.save('MoveT.xlsx')

# worksheet.move_range('B1:D3', 3)
# print(worksheet['D1'].value)
# print(worksheet['D3'].value)
# print(worksheet['D4'].internal_value)
# print(worksheet['D6'].internal_value)
# workbook.save('Move.xlsx')

# worksheet.merge_cells('C1:F1')
# worksheet.merge_cells('A1:B4')
# print('A1' in worksheet.merged_cells)
# workbook.save('Merge.xlsx')
# worksheet.merge_cells('A1')
# workbook.save('Merge2.xlsx')

# worksheet.merge_cells('B1:D3')
# worksheet['B1'].value = 'nihao'
# print(worksheet['B1'])
# print(worksheet['B2'])
# worksheet.unmerge_cells('B1:D3')
# print(worksheet['B1'])
# print(worksheet['B2'])
# workbook.save('Unmerge.xlsx')

# worksheet.insert_cols(1, 2)
# workbook.save('Insert.xlsx')

# worksheet.delete_cols(1)
# worksheet.delete_rows(1)
# workbook.save('Delete.xlsx')

# worksheet.title = 'ABC'
# print(worksheet.title)
# print(worksheet.encoding)
# print(worksheet.path)
# print(worksheet.calculate_dimension())
# print(worksheet.dimensions)
# from openpyxl.worksheet.pagebreak import RowBreak
# print(worksheet.row_breaks.append(RowBreak(4,40)))
# workbook.save('Breaks.xlsx')
# for c in worksheet.rows:
#     print(c)

# worksheet.merged_cell_ranges
# print(worksheet.protection.deleteRows)
# worksheet.protection.deleteRows = not worksheet.protection.deleteRows
# worksheet.protection.insertColumns = not worksheet.protection.insertColumns
# workbook.save('P.xlsx')

# worksheet.selected_cell
# worksheet.sheet_state = openpyxl.worksheet.worksheet.Worksheet.SHEETSTATE_HIDDEN
# worksheet.parent.create_sheet()
# workbook.save('State.xlsx')

# print(worksheet.active_cell)
# print(worksheet.array_formulae)
# print(worksheet.scenarios)
# for r in worksheet.values.__next__:
#     print(type(r))

#     for c in r:
#         print(c)

# print(worksheet.array_formulae)
print(worksheet.protection.deleteRows)
