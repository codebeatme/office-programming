from openpyxl.styles import Font, Border, Side
from openpyxl.styles.named_styles import NamedStyle

# 名稱為 My Style 的命名格式
NamedStyle(
    'My Style',
    # 字型，20 磅
    font=Font(sz=20),
    # 下邊線為紅色雙實線
    border=Border(bottom=Side('double', 'FF0000'))
)
