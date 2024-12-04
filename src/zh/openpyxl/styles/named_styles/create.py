from openpyxl.styles import Font, Border, Side
from openpyxl.styles.named_styles import NamedStyle

# 名称为 My Style 的命名格式
NamedStyle(
    'My Style',
    # 字体，20 磅
    font=Font(sz=20),
    # 下边框为红色双实线
    border=Border(bottom=Side('double', 'FF0000'))
)
