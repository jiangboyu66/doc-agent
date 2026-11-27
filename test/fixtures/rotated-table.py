"""生成测试夹具 rotated-table.pdf：表格旁边、表格里跨几行的竖排文字（从下往上读）。
  - 第一张表（画了格线）：表头那一行上方没有线，表格只认出下面三行；一行竖排的标签从表头一直竖到第三行，
    穿过两条格线。它不属于哪一格——以前被分进某一格，可编辑排版只能把它横着接在格内文字后面，那一行多出两行字高。
  - 第二张表（三线表：只有顶线、表头线、底线，行与行之间没有线）：第 3–6 行的第二、三列是两段竖排文字，
    原文是跨四行的合并格。按文字行切表时被切开，竖排文字分进其中一行，那一行按文字的长度加高，
    整张表比原文高出一大截。"""
import sys
from reportlab.pdfgen import canvas

out = sys.argv[1] if len(sys.argv) > 1 else "rotated-table.pdf"
W, H = 595, 842
c = canvas.Canvas(out, pagesize=(W, H))


def y(v):
    return H - v


def vertical(x, bottom, text, font="Helvetica", size=9):
    """从下往上读的一行竖排文字：基线在 x，起点在 bottom（页面坐标，y 向下）"""
    c.setFont(font, size)
    c.saveState()
    c.translate(x, y(bottom))
    c.rotate(90)
    c.drawString(0, 0, text)
    c.restoreState()


c.setFont("Helvetica-Bold", 14)
c.drawString(72, y(80), "Text format and page layout")

# 第一张表：表头没有上边线；格线 4 横 4 竖
c.setFont("Helvetica-Bold", 9)
c.drawString(80, y(122), "#")
c.drawString(108, y(122), "Method / Attribute")
c.drawString(330, y(122), "Description")
c.setLineWidth(0.6)
hy = [130, 152, 174, 196]
vx = [72, 100, 300, 520]
for v in hy:
    c.line(vx[0], y(v), vx[-1], y(v))
for x in vx:
    c.line(x, y(hy[0]), x, y(hy[-1]))
c.setFont("Helvetica", 9)
rows1 = [("1", "Document.pageCount", "the number of pages (int)"), ("2", "Document.metadata", "the metadata (dict)"), ("3", "Document.getToC()", "get the table of contents (list)")]
for k, (a, b, d) in enumerate(rows1):
    base = hy[k] + 14
    c.drawString(80, y(base), a)
    c.drawString(108, y(base), b)
    c.drawString(330, y(base), d)
# 竖排标签：从第三行一直竖到表头上方（穿过两条格线、伸出表格）
vertical(310, 192, "Vertical text from bottom to top", size=8)

# 第二张表：三线表
c.setFont("Helvetica-Bold", 12)
c.drawString(72, y(240), "Table border and shading style")
top, head, bottom = 255, 277, 420
c.setLineWidth(1)
c.line(72, y(top), 520, y(top))
c.setLineWidth(0.5)
c.line(72, y(head), 520, y(head))
c.setLineWidth(1)
c.line(72, y(bottom), 520, y(bottom))
c.setFont("Helvetica-Bold", 9)
c.drawString(80, y(270), "#")
c.drawString(108, y(270), "Method / Attribute")
c.drawString(330, y(270), "Description")
c.setFont("Helvetica", 9)
c.drawString(80, y(292), "1")
c.drawString(108, y(292), "Document.pageCount")
c.drawString(330, y(292), "the number of pages (int)")
c.drawString(80, y(312), "2")
c.drawString(108, y(312), "Document.metadata")
c.drawString(330, y(312), "the metadata (dict)")
for k in range(4):
    c.drawString(80, y(334 + 22 * k), str(3 + k))
# 第 3–6 行：第二、三列各一段竖排文字，跨四行
vertical(160, 412, "Document.getToC()", size=9)
vertical(172, 412, "Document.loadPage()", size=9)
for k, s in enumerate(["read a page read", "a page read a", "page read a page"]):
    vertical(360 + 12 * k, 412, s, size=9)

# 表格下面的正文
c.setFont("Helvetica", 10)
for k, s in enumerate([
    "The second table has no lines between its rows: the vertical text in the last four rows belongs to",
    "one merged cell in the original document, and the table keeps the height it has on this page.",
]):
    c.drawString(72, y(450 + 14 * k), s)
c.showPage()
c.save()
