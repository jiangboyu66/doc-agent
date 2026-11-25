"""生成测试夹具（取自 PWB 基准里失分的几类文件，文字为合成内容）：

  ragged.pdf     2 页英文协议：左对齐（右边参差）的正文排在留有内边距的文本框里，页眉横线和页脚伸到页边，
                 区域右边界比实际折行处靠右；一段跨页；页脚 "Page n of 2" 与补零的 Bates 号 ALX00000n。
                 易错点：句末恰好在行尾、或下一行以 and / of 这样的短词开头时被当成段落结束；
                 跨页的段落被切开；Bates 号逐页不同，第二页的页脚回到正文末尾。
  cells.pdf      带完整框线的付款表，标准 Times 字体：说明列的文字离右边框线很近，pdf.js 把它和右边一格的日期
                 当作词间空格合成一项。易错点：日期并进说明格，日期格变空。
  rotated.pdf    第 1 页横向（表格），第 2 页是带 /Rotate 90 的纵向正文页。易错点：整页文字被当成旋转文字丢掉，
                 或者按第一页（横向）的版心判断折行，每行都成了一段。
  footnotes.pdf  3 页备忘录，正文段末有上标引用号，各页页底有对应的脚注。易错点：只有第一页的脚注被移进首页页脚。
"""
import io
import sys
from pathlib import Path

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle

out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent
W, H = A4
SENTENCES = [
    "The Supplier warrants that the Services will be performed with reasonable skill and care in accordance with Good Industry Practice.",
    "Each Party shall keep confidential all Confidential Information disclosed to it under this Agreement and shall not use it for any purpose other than the Project.",
    "The aggregate liability of the Supplier under this Agreement shall not exceed the Charges paid in the twelve months before the claim.",
    "Late payments shall bear interest at four per cent per annum above the base rate from the due date until payment is made in full.",
    "Neither Party shall be liable for any delay caused by a Force Majeure Event lasting less than thirty days after notice is given.",
    "This Agreement constitutes the entire agreement between the Parties and supersedes all previous drafts and arrangements of any kind.",
]
PARAS = [" ".join(SENTENCES[(i + k) % len(SENTENCES)] for k in range(3)) for i in range(12)]


def ragged() -> None:
    body = ParagraphStyle("b", fontName="Times-Roman", fontSize=10.5, leading=15, spaceAfter=8)

    def page(c, d):
        c.saveState()
        c.setFont("Times-Roman", 8)
        c.drawString(25 * mm, H - 14 * mm, "Project Kestrel - Services Agreement")
        c.line(25 * mm, H - 16 * mm, W - 25 * mm, H - 16 * mm)
        c.drawCentredString(W / 2, 12 * mm, f"Page {d.page} of 2")
        c.drawRightString(W - 25 * mm, 12 * mm, f"ALX{d.page:06d}")
        c.restoreState()

    d = SimpleDocTemplate(str(out / "ragged.pdf"), pagesize=A4, leftMargin=25 * mm, rightMargin=25 * mm, topMargin=25 * mm, bottomMargin=25 * mm)
    d.build([Paragraph(p, body) for p in PARAS[:12]], onFirstPage=page, onLaterPages=page)


def cells() -> None:
    rows = [["Milestone", "Description", "Due date", "Amount (USD)"]] + [
        [f"M{i:02d}", f"Delivery of work package {i} and acceptance testing", f"2027-0{i}-14", f"{300 + 10 * i},000"] for i in range(1, 6)]
    t = Table(rows, colWidths=[20 * mm, 62 * mm, 26 * mm, 30 * mm])
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.4, "black"), ("FONT", (0, 0), (-1, -1), "Times-Roman", 8.5)]))
    SimpleDocTemplate(str(out / "cells.pdf"), pagesize=A4).build([t])


def rotated() -> None:
    from pypdf import PdfReader, PdfWriter

    first = io.BytesIO()
    t = Table([["Item", "Specification", "Qty"]] + [[str(i), f"Industrial valve type V-{100 + i} rated PN16", str(i * 3)] for i in range(1, 6)])
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.4, "black"), ("FONT", (0, 0), (-1, -1), "Times-Roman", 9)]))
    SimpleDocTemplate(first, pagesize=landscape(A4)).build([t])
    second = io.BytesIO()
    body = ParagraphStyle("b", fontName="Times-Roman", fontSize=10.5, leading=15, spaceAfter=8)
    SimpleDocTemplate(second, pagesize=A4).build([Paragraph(p, body) for p in PARAS[:4]])
    w = PdfWriter()
    w.add_page(PdfReader(first).pages[0])
    page = PdfReader(second).pages[0]
    page.rotate(90)
    w.add_page(page)
    with open(out / "rotated.pdf", "wb") as f:
        w.write(f)


def footnotes() -> None:
    c = canvas.Canvas(str(out / "footnotes.pdf"), pagesize=A4)
    note = 0
    for pg in range(3):
        y = H - 30 * mm
        if pg == 0:
            c.setFont("Times-Bold", 14)
            c.drawString(25 * mm, y, "MEMORANDUM")
            y -= 12 * mm
        notes = []
        for k in range(3):
            text = SENTENCES[(pg + k) % len(SENTENCES)]
            note += 1
            c.setFont("Times-Roman", 10.5)
            c.drawString(25 * mm, y, text[:90])
            y -= 15
            last = text[90:] or "See below."
            c.drawString(25 * mm, y, last)
            c.setFont("Times-Roman", 6.5)
            c.drawString(25 * mm + pdfmetrics.stringWidth(last, "Times-Roman", 10.5) + 1, y + 4, str(note))
            y -= 23
            notes.append(f"{note} See clause {note + 2}.1 of the {['Framework', 'Supply', 'Escrow'][note % 3]} Agreement dated {note + 9} March 2026.")
        c.line(25 * mm, 32 * mm, 80 * mm, 32 * mm)
        c.setFont("Times-Roman", 7.5)
        yy = 28 * mm
        for n in notes:
            c.drawString(25 * mm, yy, n)
            yy -= 10
        c.showPage()
    c.save()


ragged()
cells()
rotated()
footnotes()
