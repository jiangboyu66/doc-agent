"""生成测试夹具 mixed-pages.pdf：一份文件里纸张大小 / 方向 / 页面框不一样的页。
  第 1 页：纵向 Letter 正文；
  第 2 页：横向 Letter，一张 7 列、横跨整个横向版面的宽表格（Word 文档的横向附录）；
  第 3 页：纵向 A4（Excel 导出的说明页）；
  第 4 页：/Rotate 180，内容按未旋转的方向排（扫描、合并时转了页面方向）；
  第 5 页：CropBox 比 MediaBox 小、有偏移，裁剪框外面还画了字和色块（看不见，不该转换）。
以前整篇都用第一页的尺寸：横向页被转成纵向、表格最右一列被裁掉；裁剪框外的东西按负坐标排进版面，一页变几页。"""
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color

out = sys.argv[1] if len(sys.argv) > 1 else "mixed-pages.pdf"
LETTER = (612, 792)
c = canvas.Canvas(out, pagesize=LETTER)


def para(x, y, lines, size=11, font="Helvetica"):
    c.setFont(font, size)
    for i, s in enumerate(lines):
        c.drawString(x, y - i * size * 1.3, s)


# 第 1 页：纵向正文
W, H = LETTER
para(72, H - 90, ["Scheduler service design"], 16, "Helvetica-Bold")
para(72, H - 120, [
    "This document describes the job scheduler. Jobs are submitted through the API and stored",
    "in the queue until a worker picks them up. The event schema is listed in Appendix A, which",
    "is printed on a landscape page because the table is wider than a portrait page.",
])
c.showPage()

# 第 2 页：横向，宽表格
W, H = 792, 612
c.setPageSize((W, H))
para(54, H - 60, ["Appendix A. Event schema (landscape)"], 14, "Helvetica-Bold")
cols = [90, 80, 70, 70, 120, 110, 144]
heads = ["Field", "Type", "Required", "Default", "Example", "Owner", "Description"]
rows = [
    ["job_id", "UUID", "yes", "-", "3f2a91c", "api", "Stable identifier"],
    ["tenant", "string", "yes", "-", "acme-eu", "api", "Owning tenant"],
    ["run_at", "timestamp", "yes", "-", "2026-03-01", "scheduler", "UTC start time"],
    ["priority", "int8", "no", "5", "1", "scheduler", "Lowest is nine"],
    ["retries", "int8", "no", "3", "5", "worker", "Attempts before giving up"],
]
x0, top = 54, H - 90
xs = [x0]
for w in cols:
    xs.append(xs[-1] + w)
ys = [top - 22 * i for i in range(len(rows) + 2)]
c.setLineWidth(0.6)
for xv in xs:
    c.line(xv, ys[0], xv, ys[-1])
for yv in ys:
    c.line(xs[0], yv, xs[-1], yv)
for r, row in enumerate([heads] + rows):
    c.setFont("Helvetica-Bold" if r == 0 else "Helvetica", 10)
    for k, v in enumerate(row):
        c.drawString(xs[k] + 4, ys[r] - 15, v)
c.showPage()

# 第 3 页：纵向 A4
W, H = 595, 842
c.setPageSize((W, H))
para(60, H - 80, ["Notes"], 14, "Helvetica-Bold")
para(60, H - 105, ["Figures are preliminary and unaudited. Subtotals exclude hidden rows."])
c.showPage()

# 第 4 页：/Rotate 180，内容按未旋转的方向排
W, H = LETTER
c.setPageSize(LETTER)
c.setPageRotation(180)
para(72, H - 90, ["Rotated page heading"], 14, "Helvetica-Bold")
para(72, H - 115, ["This page carries a rotation of one hundred eighty degrees in its page dictionary."])
c.showPage()

# 第 5 页：CropBox 有偏移，裁剪框外面还有字和色块
c.setPageSize(LETTER)
c.setPageRotation(0)
c.setCropBox((100, 150, 500, 650))
para(130, 600, ["Visible text inside the crop box."])
para(130, 580, ["The crop box hides everything outside it."])
c.setFillColor(Color(1, 0.6, 0.6))
c.rect(20, 700, 60, 40, stroke=0, fill=1)
c.setFillColor(Color(0, 0, 0))
para(20, 60, ["Hidden text below the crop box."])
para(520, 400, ["Hidden right"])
c.showPage()
c.save()
