# 請將命令列跳躍至 Goods.xlsx 所在的目錄，然後執行此腳本檔案
from openpyxl import load_workbook

workbook = load_workbook('Goods.xlsx')
worksheet = workbook.create_sheet()
# 以不同形式為工作表新增 4 列資料
worksheet.append((1.6, 2.5))
# 這相當於將儲存格 A1 和 B1 移動至 A2 和 B2 的所在位置
worksheet.append([worksheet['A1'], worksheet['B1']])
worksheet.append({2: 1.0, 4: 2.0})
worksheet.append({'C': 1.1, 'E': 2.4})
# 儲存為 Append.xlsx
workbook.save('Append.xlsx')

# 顯示所有儲存格的值
for r in worksheet.values:
    print(r)


from openpyxl import Workbook

workbook = Workbook(True)
worksheet = workbook.create_sheet()
# 為唯寫工作表新增資料，並儲存為 AppendWriteOnly.xlsx
worksheet.append((1.6, 2.5))
worksheet.append([1.8, 2.2])
workbook.save('AppendWriteOnly.xlsx')
