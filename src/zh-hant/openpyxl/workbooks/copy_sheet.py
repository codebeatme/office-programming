# 請將命令列跳躍至 School.xlsx 所在的目錄，然後執行此腳本檔案
import openpyxl
workbook = openpyxl.load_workbook('School.xlsx')

# 複製工作表 ClassA
workbook.copy_worksheet(workbook['ClassA'])
print(f'ClassA 複本的名稱為 {workbook.worksheets[-1].title}')

# 嘗試複製其他活頁簿的工作表
newbook = openpyxl.Workbook()
# ERROR 無法複製自身不擁有的工作表
workbook.copy_worksheet(newbook['Sheet'])
