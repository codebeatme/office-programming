# 请将命令行跳转至 School.xlsx 所在的目录，然后运行此脚本文件
import openpyxl
workbook = openpyxl.load_workbook('School.xlsx')

# 复制工作表 ClassA
workbook.copy_worksheet(workbook['ClassA'])
print(f'ClassA 副本的名称为 {workbook.worksheets[-1].title}')

# 尝试复制其他工作簿的工作表
newbook = openpyxl.Workbook()
# ERROR 无法复制自身不拥有的工作表
workbook.copy_worksheet(newbook['Sheet'])
