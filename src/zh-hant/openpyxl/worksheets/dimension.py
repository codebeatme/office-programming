# 請將命令列跳躍至 Goods.xlsx 所在的目錄，然後執行此腳本檔案
from openpyxl import load_workbook

empty = load_workbook('Goods.xlsx')['Empty']
# C3 是一個值為空但擁有背景色彩的儲存格
print(empty['C3'])
print(f'最小範圍 {empty.dimensions}')
# 存取 E5 導致最小範圍改變
empty['E5']
print(f'存取 E5 後的最小範圍 {empty.calculate_dimension()}')

r_empty = load_workbook('Goods.xlsx', True)['Empty']
# 對於唯讀工作表，存取 E5 不會導致最小範圍改變
r_empty['E5']
print(f'唯讀工作表存取 E5 後的最小範圍 {r_empty.calculate_dimension()}')

# 重設最大列和最大欄
r_empty.reset_dimensions()
# ERROR 需要將 force 參數設定為 True
r_empty.calculate_dimension()