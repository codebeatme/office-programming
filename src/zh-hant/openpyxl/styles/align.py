from openpyxl.styles import Alignment

# 水平置中，垂直置中，自動換行
Alignment('center', 'center', wrapText=True)
# 需要時縮小文字，行首縮排大小 1
Alignment(shrinkToFit=True, indent=1)
# 文字方向從右到左
Alignment(readingOrder=2)