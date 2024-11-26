from openpyxl.styles import Alignment

# 水平居中，垂直居中，自动换行
Alignment('center', 'center', wrapText=True)
# 需要时缩小文字，行首缩进大小 1
Alignment(shrinkToFit=True, indent=1)
# 文本方向从右到左
Alignment(readingOrder=2)