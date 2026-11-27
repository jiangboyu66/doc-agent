"""生成测试夹具 dense-cells.pdf：选区公报式的密排表格（横向）。表头两行，下面一行很高，三格里各是一整栏
"候选人  党派  票数"，票数右对齐、离右框线只有 2pt（比 Word 单元格的默认边距还近）。
以前可编辑排版里每个票数都折到下一行，整格高出一倍；这一行又不许跨页，整行挤到第二页，第一页只剩表头。"""
import sys
from reportlab.pdfgen import canvas

out = sys.argv[1] if len(sys.argv) > 1 else "dense-cells.pdf"
W, H = 792, 612
c = canvas.Canvas(out, pagesize=(W, H))


def y(v):
    return H - v


xs = [28.8, 270, 511, 763.2]
ys = [20, 40, 58, 590]
c.setLineWidth(0.5)
for v in ys:
    c.line(xs[0], y(v), xs[-1], y(v))
for x in (xs[0], xs[-1]):
    c.line(x, y(ys[0]), x, y(ys[-1]))
for x in xs[1:-1]:
    c.line(x, y(ys[2]), x, y(ys[-1]))
c.setFont("Helvetica", 8)
c.drawString(32, y(33), "PRECINCT BULLETIN   GENERAL ELECTION   NOVEMBER 4")
c.drawRightString(760, y(33), "PAGE 1")
c.drawString(32, y(52), "327 BALLOTS CAST   1100 REGISTERED VOTERS")
names = ["ADAMS", "BAKER", "CLARK", "DAVIS", "EVANS", "FOSTER", "GARCIA", "HUGHES", "IRWIN", "JONES", "KELLY", "LOPEZ"]
for col in range(3):
    x0, x1 = xs[col], xs[col + 1]
    line = 70
    for race in range(5):
        c.drawString(x0 + 4, y(line), f"CONTEST {col * 5 + race + 1}")
        c.drawString(x0 + 4, y(line + 10), "VOTER NOMINATED")
        line += 20
        for k in range(8):
            c.drawString(x0 + 4, y(line), f"{names[(k + race + col) % 12]} {chr(65 + k)}. SMITH")
            c.drawString(x0 + 150, y(line), ["REP", "DEM"][k % 2])
            c.drawRightString(x1 - 2, y(line), str(37 * (k + 1) + 11 * race + col))
            line += 10
        line += 6
c.showPage()
c.save()
