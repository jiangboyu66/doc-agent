"""生成测试夹具 overprint.pdf：没有粗体字形的字体"叠印加粗"——表头的每个词在几乎同一位置画了四遍
（每遍错开 0.2pt）。以前每一遍都是一段独立的文字：可编辑排版的表头格里同一个词排成四遍，整行撑高。"""
import sys
from reportlab.pdfgen import canvas

out = sys.argv[1] if len(sys.argv) > 1 else "overprint.pdf"
W, H = 595, 842
c = canvas.Canvas(out, pagesize=(W, H))


def y(v):
    return H - v


def bold(x, base, text, size=10):
    """叠印加粗：同一段文字画四遍，每遍错开 0.2pt"""
    c.setFont("Helvetica", size)
    for dx, dy in [(0, 0), (0.2, 0), (0, 0.2), (0.2, 0.2)]:
        c.drawString(x + dx, y(base) + dy, text)


bold(72, 80, "Insurance plan summary", 14)
xs = [72, 172, 302, 412, 523]
ys = [100, 124, 146, 168, 190]
c.setLineWidth(0.6)
for v in ys:
    c.line(xs[0], y(v), xs[-1], y(v))
for x in xs:
    c.line(x, y(ys[0]), x, y(ys[-1]))
for k, h in enumerate(["Plan", "Term", "Entry age", "Premium"]):
    bold(xs[k] + 6, 116, h)
c.setFont("Helvetica", 10)
rows = [["Basic", "10 years", "15 - 60", "Monthly"], ["Standard", "15 years", "15 - 55", "Monthly"], ["Premium", "20 years", "15 - 50", "Yearly"]]
for r, row in enumerate(rows):
    for k, v in enumerate(row):
        c.drawString(xs[k] + 6, y(ys[r + 1] + 15), v)
c.drawString(72, y(220), "Header words above are drawn four times each, offset by a fifth of a point.")
c.showPage()
c.save()
