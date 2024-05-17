# overwrite.xlsx 会被保存至命令行的当前工作目录
# 导入包 openpyxl
import openpyxl

# 创建 Workbook 并先后保存两次
workbook = openpyxl.Workbook()
workbook['Sheet']['A1'].value = '世界'
workbook.save('overwrite.xlsx')
workbook['Sheet']['A1'].value = '你好！'
workbook.save('overwrite.xlsx')
