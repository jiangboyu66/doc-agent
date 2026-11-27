"""生成测试夹具 open-table.pdf：左右两边不画竖线的统计表（只有内部竖线，横线左右伸到表格边上）。
第一列"行业名称"两行合并、最后一列"微型"有空格子。以前表格只按竖线算宽：最左、最右两列落在表格外，
被另一种表格识别成一张两列的表，中间各列成了另一张表排到后面，一页变两页。"""
import sys
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

out = sys.argv[1] if len(sys.argv) > 1 else "open-table.pdf"
pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
W, H = 595, 842
c = canvas.Canvas(out, pagesize=(W, H))
F = "STSong-Light"

c.setFont(F, 14)
c.drawCentredString(W / 2, H - 70, "统计上大中小微型企业划分标准")

cols = [62, 168, 236, 271, 327, 412, 484, 534]  # 表格左右边界 62 / 534 处没有竖线
heads = ["行业名称", "指标名称", "计量单位", "大型", "中型", "小型", "微型"]
inds = [
    ("农、林、牧、渔业", [("营业收入(Y)", "万元", "Y≥20000", "500≤Y<20000", "50≤Y<500", "Y<50")]),
    ("工业", [("从业人员(X)", "人", "X≥1000", "300≤X<1000", "20≤X<300", "X<20"), ("营业收入(Y)", "万元", "Y≥40000", "2000≤Y<40000", "300≤Y<2000", "Y<300")]),
    ("建筑业", [("营业收入(Y)", "万元", "Y≥80000", "6000≤Y<80000", "300≤Y<6000", "Y<300"), ("资产总额(Z)", "万元", "Z≥80000", "5000≤Z<80000", "300≤Z<5000", "")]),
    ("批发业", [("从业人员(X)", "人", "X≥200", "20≤X<200", "5≤X<20", "X<5"), ("营业收入(Y)", "万元", "Y≥40000", "5000≤Y<40000", "1000≤Y<5000", "Y<1000")]),
    ("零售业", [("从业人员(X)", "人", "X≥300", "50≤X<300", "10≤X<50", ""), ("营业收入(Y)", "万元", "Y≥20000", "500≤Y<20000", "100≤Y<500", "Y<100")]),
    ("交通运输业", [("从业人员(X)", "人", "X≥1000", "300≤X<1000", "20≤X<300", "X<20"), ("营业收入(Y)", "万元", "Y≥30000", "3000≤Y<30000", "200≤Y<3000", "Y<200")]),
]
row_h = 20
top = H - 95
c.setLineWidth(0.48)
y = top
hlines = [y]
y -= 30
hlines.append(y)
rows = []
for name, rs in inds:
    start = y
    for r in rs:
        rows.append((y, r))
        y -= row_h
    rows_name = (start, y, name)
    rows.append(("name", rows_name))
    hlines.append(y)
bottom = y
# 横线：按列分段画，左右伸到表格边上
for yy in hlines:
    for k in range(len(cols) - 1):
        c.line(cols[k] + (0.5 if k else 0), yy, cols[k + 1], yy)
# 竖线：只有内部的
for x in cols[1:-1]:
    c.line(x, top, x, bottom)
c.setFont(F, 9)
for k, h in enumerate(heads):
    c.drawCentredString((cols[k] + cols[k + 1]) / 2, top - 19, h)
for it in rows:
    if it[0] == "name":
        s, e, name = it[1]
        c.drawCentredString((cols[0] + cols[1]) / 2, (s + e) / 2 - 3, name)
        continue
    yy, r = it
    for k, v in enumerate(r):
        if v:
            c.drawCentredString((cols[k + 1] + cols[k + 2]) / 2, yy - 13.5, v)
c.setFont(F, 10.5)
c.drawString(62, bottom - 30, "说明：大型、中型和小型企业须同时满足所列指标的下限，否则下划一档。")
c.showPage()
c.save()
