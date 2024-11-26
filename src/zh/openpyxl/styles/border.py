from openpyxl.styles import Border, Side, Color

# 左右为红色双实线，上下为蓝色粗实线的边框
double = Side('double', 'FF0000')
thick = Side('thick', Color('0000FF'))
Border(double, double, thick, thick)