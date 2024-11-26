from openpyxl.styles import Border, Side, Color

# 左右為紅色雙實線，上下為藍色粗實線的邊線
double = Side('double', 'FF0000')
thick = Side('thick', Color('0000FF'))
Border(double, double, thick, thick)