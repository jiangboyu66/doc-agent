"""生成测试夹具 rotated-text.pdf：页面上的旋转文字。
  - 带框表格的表头里有两格是竖排的（从下往上读），其余是横排；
  - 页面左边距里有一行竖排的说明（从下往上读，期刊、政府公报常见）；
  - 正文旁边有一个从上往下读的竖排标签；
  - 一个斜放 35° 的半透明水印 "CONFIDENTIAL"，字距加宽（pdf.js 会在字母之间插空格）。
以前旋转文字整段不转换，报告里却显示文字完整性 100%。"""
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color

out = sys.argv[1] if len(sys.argv) > 1 else "rotated-text.pdf"
W, H = 595, 842
c = canvas.Canvas(out, pagesize=(W, H))
black = Color(0, 0, 0)


def y(v):
    return H - v


# 正文
c.setFont("Helvetica-Bold", 16)
c.drawString(72, y(90), "Quarterly Sales Summary")
c.setFont("Helvetica", 10)
body = [
    "This report summarises the quarterly sales of the three regional offices. All figures are",
    "in thousands of euros and exclude value added tax. The table below lists each office with",
    "its sales in the four quarters of the year; the column headers of the quarters are printed",
    "vertically so that the table fits the width of the page.",
]
for i, s in enumerate(body):
    c.drawString(72, y(120 + 14 * i), s)

# 带框表格：第一列是办公室，后面四列的表头竖排
x0, top = 72, 190
cols = [150, 60, 60, 60, 60]
rows = [70, 20, 20, 20]
xs = [x0]
for w in cols:
    xs.append(xs[-1] + w)
ys = [top]
for h in rows:
    ys.append(ys[-1] + h)
c.setLineWidth(0.8)
for xv in xs:
    c.line(xv, y(ys[0]), xv, y(ys[-1]))
for yv in ys:
    c.line(xs[0], y(yv), xs[-1], y(yv))
c.setFont("Helvetica-Bold", 10)
c.drawString(xs[0] + 4, y(ys[1] - 6), "Office")
for k, label in enumerate(["Quarter one", "Quarter two", "Quarter three", "Quarter four"]):
    cx = xs[k + 1] + cols[k + 1] / 2 + 3.5
    c.saveState()
    c.translate(cx, y(ys[1] - 5))
    c.rotate(90)  # 从下往上读
    c.drawString(0, 0, label)
    c.restoreState()
c.setFont("Helvetica", 10)
data = [["Berlin", "120", "135", "128", "142"], ["Madrid", "98", "104", "111", "119"], ["Vienna", "87", "90", "93", "101"]]
for r, row in enumerate(data):
    base = ys[r + 2] - 6
    c.drawString(xs[0] + 4, y(base), row[0])
    for k, v in enumerate(row[1:]):
        c.drawRightString(xs[k + 2] - 6, y(base), v)

# 正文旁边从上往下读的竖排标签
c.setFont("Helvetica-Oblique", 9)
c.saveState()
c.translate(520, y(400))
c.rotate(-90)
c.drawString(0, 0, "Figures unaudited")
c.restoreState()
c.setFont("Helvetica", 10)
more = [
    "Sales grew in every office during the year. The strongest growth was in Berlin, where a new",
    "distribution partner started in the second quarter. Madrid and Vienna grew steadily without",
    "any special effects. The outlook for the coming year is positive for all three offices.",
]
for i, s in enumerate(more):
    c.drawString(72, y(400 + 14 * i), s)

# 左边距里从下往上读的说明
c.setFont("Helvetica", 7)
c.saveState()
c.translate(36, y(700))
c.rotate(90)
c.drawString(0, 0, "Internal draft prepared by the finance office")
c.restoreState()

# 斜放的半透明水印，字距加宽
c.saveState()
c.setFillColor(Color(0.8, 0, 0))
c.setFillAlpha(0.15)
c.translate(150, y(650))
c.rotate(35)
t = c.beginText(0, 0)
t.setFont("Helvetica-Bold", 60)
t.setCharSpace(12)
t.textOut("CONFIDENTIAL")
c.drawText(t)
c.restoreState()

c.showPage()
c.save()
