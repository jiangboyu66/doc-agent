"""生成测试夹具 three-columns.pdf：通栏标题 + 三栏正文（报纸、公报式），两页。
以前只找一条栏间距：三栏被当成两栏，其中一"栏"里并排着两栏的文字，左右两栏的行被配成一行，交错重叠，页数翻倍。"""
import sys
from reportlab.pdfgen import canvas

out = sys.argv[1] if len(sys.argv) > 1 else "three-columns.pdf"
W, H = 612, 792
c = canvas.Canvas(out, pagesize=(W, H))
COLS = [(54, 214), (226, 386), (398, 558)]
FONT, SIZE, LEAD = "Times-Roman", 9.5, 11.5
words = ("The rule proposes new inspection requirements for the flight control systems of the affected airplanes and asks for "
         "comments on the costs of compliance, the number of airplanes registered in the country, and the time needed to "
         "complete the work at the next scheduled maintenance visit ").split()


def y(v):
    return H - v


def column(x0, x1, top, bottom, lead_words, k0):
    """一栏正文：开头几个词标明是第几栏，之后逐词折行，排到 bottom 为止"""
    seq = lead_words + [words[(k0 + i) % len(words)] for i in range(2000)]
    line, base, k = "", top, 0
    while base <= bottom:
        w = seq[k]
        k += 1
        t = (line + " " + w).strip()
        if c.stringWidth(t, FONT, SIZE) > x1 - x0:
            c.setFont(FONT, SIZE)
            c.drawString(x0, y(base), line)
            base += LEAD
            line = w
        else:
            line = t


for page in range(2):
    c.setFont("Times-Bold", 9)
    c.drawString(54, y(40), "Federal Notices")
    c.drawRightString(558, y(40), f"Page {page + 1}")
    c.setLineWidth(0.5)
    c.line(54, y(46), 558, y(46))
    top = 70
    if page == 0:
        c.setFont("Times-Bold", 16)
        c.drawString(54, y(78), "Airworthiness Directives; Transport Category Airplanes")
        top = 104
    names = ["One", "Two", "Three"]
    for k, (x0, x1) in enumerate(COLS):
        column(x0, x1, top, 740, f"Column {names[k]} of page {page + 1} begins here.".split(), 7 * k + 3 * page)
    c.showPage()
c.save()
