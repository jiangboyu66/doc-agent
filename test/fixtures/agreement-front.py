"""生成测试夹具 agreement-front.pdf：美式资产购买协议的前几页（虚构公司），8 页，Letter。
  第 1 页：封面（居中的标题行），左下角小字号的文档编号"[[9876543]]"（每页都有）。
  第 2–3 页：目录。每条一行："Section 1.1. Definitions……1"，引导点是文字、页码顶到右边界；
          有一条标题折成两行，只有第二行带引导点和页码；各章之间空一行。页脚居中印"~i~""~ii~"（罗马数字页码）。
  第 4 页：附表清单（"SCHEDULE LIST"，下面一张有框线的两列表），页面下半部空着，页脚"~iii~"。
  第 5 页：附件清单（"EXHIBITS"，标题下面一条下划线），页脚"~iv~"。
  第 6–8 页：正文，页脚居中印阿拉伯数字页码 1、2、3。定义条款"“Accounts Payable” shall mean …"一整行一次画出
          （pdf.js 读成一个文字项），只有被定义的词下面有一条下划线（扁的填充矩形）；前言里的 (“Northwind”)、
          行中间的 (the “Closing Date”) 也一样。"W I T N E S S E T H :" 逐字母加宽排，每个字母下面一小段线。
          第 7–8 页的"定义术语索引"：左边术语、右边所在条款（右栏从版心 80% 处起排），第 8 页整页都是索引。
易错点（2026-11-27 APA 英译中，pdf2word 1.1.3）：
  - 目录每条的行都排满到右边界，被当成一段正文的折行，整章的条目并成一段，引导点在译文里丢掉，页码贴在标题后面；
  - "~i~"不是数字，认不出是页码，留在正文里靠段前距排到页底：译文变短后排到页面中间；它占着页底，
    附表清单、附件清单、协议正文都没有另起一页；
  - 下划线盖不到文字项的一半，被当成分隔线，写成行间一个带下边框的空段落：定义条款从第一行之后断成两段，
    译文换了字宽以后线与文字错开（APA 357 条，26 条定义的译文被切成两半、中间补了"……"）；
  - 定义术语索引的右栏起点超出版心 60%，不算释义清单，整页被当成分栏，术语与条款号对不上。
标准答案（两种排版方式）：目录一条一段、页码前是右对齐带点状前导符的制表位；罗马数字页码写进前置部分那一节的页脚
  （PAGE 域，小写罗马数字），正文另起一节从 1 编页；"~i~"不在正文里；附表清单、附件清单另起一页；
  下划线落在被定义的词上（字符下划线），没有带边框的空段落，定义条款是一整段；索引写成两列的无边框表格。"""
import sys
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

W, H = 612.0, 792.0
L, R = 72.0, 540.0
out = sys.argv[1] if len(sys.argv) > 1 else "agreement-front.pdf"
c = canvas.Canvas(out, pagesize=(W, H))
F, FB = "Times-Roman", "Times-Bold"
PITCH = 13.8


def y(top_down: float) -> float:
    """以页面顶端为原点的基线位置 → PDF 坐标。"""
    return H - top_down


def footer(label: str) -> None:
    if label:
        c.setFont(F, 12)
        c.drawCentredString(W / 2, y(746.5), label)
    c.setFont(F, 7)
    c.drawString(L, y(755.0), "[[9876543]]")


def underline(x0: float, x1: float, base: float) -> None:
    """文字下方 1.3pt 处一条 0.6pt 高的填充矩形（Word 打印成 PDF 时的下划线）。"""
    c.rect(x0, y(base) - 1.9, x1 - x0, 0.6, stroke=0, fill=1)


def wrap(text: str, width: float, first_indent: float = 0.0, font: str = F, size: float = 12) -> list[str]:
    lines, line, avail = [], "", width - first_indent
    for word in text.split():
        trial = (line + " " + word).strip()
        if stringWidth(trial, font, size) > avail and line:
            lines.append(line)
            line, avail = word, width
        else:
            line = trial
    if line:
        lines.append(line)
    return lines


def paragraph(text: str, base: float, indent: float = 36.0, terms: tuple[str, ...] = ()) -> float:
    """首行缩进的段落；terms 里的词在所在行下面画下划线（与整行一次画出）。返回下一行的基线。"""
    c.setFont(F, 12)
    for k, line in enumerate(wrap(text, R - L, indent)):
        x = L + (indent if k == 0 else 0)
        c.drawString(x, y(base), line)
        for term in terms:
            # 加引号的定义处（"(“Northwind”)"）优先，没有时才是普通出现
            at = line.find("“" + term + "”")
            at = at + 1 if at >= 0 else line.find(term)
            if at >= 0:
                x0 = x + stringWidth(line[:at], F, 12)
                underline(x0, x0 + stringWidth(term, F, 12), base)
        base += PITCH
    return base


# ---- 第 1 页：封面 ----
c.setFont(FB, 10)
c.drawString(470, y(60), "EXECUTION VERSION")
c.setFont(FB, 16)
for k, t in enumerate(["AGREEMENT", "OF", "PURCHASE AND SALE", "by and between", "NORTHWIND HOLDINGS INC.", "and", "BLUE HERON PARTNERS, LLC"]):
    c.drawCentredString(W / 2, y(250 + 38 * k), t)
footer("")
c.showPage()

# ---- 第 2–3 页：目录 ----
toc = [
    ("ARTICLE 1. DEFINITIONS AND TERMS", "1"), ("Section 1.1. Definitions", "1"), ("Section 1.2. Index of Defined Terms.", "2"),
    ("Section 1.3. Other Definitional Provisions and Construction", "3"), None,
    ("ARTICLE 2. PURCHASE AND SALE", "4"), ("Section 2.1. Purchase and Sale of Conveyed Assets.", "4"), ("Section 2.2. Conveyed Assets", "4"),
    ("Section 2.3. Consents and Shared Contracts", "6"), ("Section 2.4. Excluded Assets of the Business", "7"),
    ("Section 2.5. Assumed Liabilities of the Business.", "8"), ("Section 2.6. Retained Liabilities of the Business", "10"),
    ("Section 2.7. Purchase Price", "11"), ("Section 2.8. Purchase Price Adjustment for Net Working Capital.", "11"), None,
    ("ARTICLE 3. CLOSING", "13"), ("Section 3.1. Closing.", "13"), ("Section 3.2. Closing; Deliveries", "14"), None,
    ("ARTICLE 4. CONDITIONS TO CLOSING", "15"), ("Section 4.1. Conditions to the Obligations of Buyer and Northwind.", "15"),
    ("Section 4.2. Conditions to the Obligations of Buyer.", "15"), ("Section 4.3. Conditions to the Obligations of Northwind", "16"),
    ("Section 4.4. Frustration of Closing Conditions", "16"), None,
    ("ARTICLE 5. REPRESENTATIONS AND WARRANTIES OF NORTHWIND", "17"), ("Section 5.1. Organization", "17"),
    ("Section 5.2. Authority; Binding Effect.", "17"), ("Section 5.3. Non-Contravention.", "18"), ("Section 5.4. Financial Information", "18"),
    ("Section 5.5. Absence of Material Adverse Effect", "19"), ("Section 5.6. Accounts Receivable", "19"),
    # 第 3 页
    ("Section 5.7. Real Property.", "20"), ("Section 5.8. Material Contracts", "20"), ("Section 5.9. Intellectual Property Rights.", "21"), None,
    ("ARTICLE 6. CONDUCT OF ACTIONS RELATING TO ENVIRONMENTAL LOSSES\nAND OBLIGATIONS RELATING THERETO", "22"),
    ("Section 6.1. Standard of Performance", "22"), ("Section 6.2. Limitations.", "22"), ("Section 6.3. Cooperation", "23"), None,
    ("ARTICLE 7. MISCELLANEOUS", "24"), ("Section 7.1. Notices.", "24"), ("Section 7.2. Amendments; Waivers.", "25"),
    ("Section 7.3. Governing Law; Jurisdiction", "25"),
]


def toc_line(title: str, page: str, base: float) -> None:
    c.setFont(F, 12)
    pw = stringWidth(page, F, 12)
    dots = "." * int((R - pw - 1 - L - stringWidth(title, F, 12)) / stringWidth(".", F, 12))
    c.drawString(L, y(base), title + dots)
    c.drawRightString(R, y(base), page)


roman = iter(["~i~", "~ii~"])
c.setFont(F, 12)
c.drawCentredString(W / 2, y(110), "TABLE OF CONTENTS")
underline(W / 2 - stringWidth("TABLE OF CONTENTS", F, 12) / 2, W / 2 + stringWidth("TABLE OF CONTENTS", F, 12) / 2, 110)
base = 142.0
for k, entry in enumerate(toc):
    if k == 31:  # 第 2 页排满，换页
        footer(next(roman))
        c.showPage()
        base = 96.0
    if entry is None:
        base += 6.0
        continue
    title, page = entry
    parts = title.split("\n")
    for extra in parts[:-1]:
        c.setFont(F, 12)
        c.drawString(L, y(base), extra)
        base += PITCH
    toc_line(parts[-1], page, base)
    base += 19.8
footer(next(roman))
c.showPage()

# ---- 第 4 页：附表清单 ----
c.setFont(FB, 12)
c.drawCentredString(W / 2, y(97), "SCHEDULE LIST")
schedules = [("Schedule 1.1(a)", "Asset Selling Corporations"), ("Schedule 1.1(b)", "Additional Employees"), ("Schedule 1.1(c)", "Knowledge Persons"),
             ("Schedule 1.1(d)", "Retained Product Liability"), ("Schedule 2.2(a)", "Owned Real Property"), ("Schedule 2.4(xi)", "Excluded Assets"),
             ("Schedule 2.8(a)", "Closing Accounts Principles"), ("Schedule 3.2(a)", "Closing Deliveries")]
c.setLineWidth(0.48)
top_row = 112.0
xs = (66.4, 198.0, 509.6)
for k, (a, b) in enumerate(schedules):
    t = top_row + 14.4 * k
    c.setFont(F, 12)
    c.drawString(72, y(t + 11.5), a)
    c.drawString(204, y(t + 11.5), b)
    c.line(xs[0], y(t), xs[2], y(t))
c.line(xs[0], y(top_row + 14.4 * len(schedules)), xs[2], y(top_row + 14.4 * len(schedules)))
for x in xs:
    c.line(x, y(top_row), x, y(top_row + 14.4 * len(schedules)))
footer("~iii~")
c.showPage()

# ---- 第 5 页：附件清单 ----
c.setFont(FB, 12)
c.drawCentredString(W / 2, y(111), "EXHIBITS")
underline(W / 2 - stringWidth("EXHIBITS", FB, 12) / 2, W / 2 + stringWidth("EXHIBITS", FB, 12) / 2, 111)
exhibits = [("Exhibit A", "Northwind Warehousing and Services Agreement"), ("Exhibit B", "Buyer Warehousing and Services Agreement"),
            ("Exhibit C", "Distributor Agreement"), ("Exhibit D", "Transition Services Agreement")]
for k, (a, b) in enumerate(exhibits):
    t = 126.4 + 14.4 * k
    c.setFont(F, 12)
    c.drawString(72, y(t + 11.5), a)
    c.drawString(204, y(t + 11.5), b)
    c.line(xs[0], y(t), xs[2], y(t))
c.line(xs[0], y(126.4 + 14.4 * len(exhibits)), xs[2], y(126.4 + 14.4 * len(exhibits)))
for x in xs:
    c.line(x, y(126.4), x, y(126.4 + 14.4 * len(exhibits)))
footer("~iv~")
c.showPage()

# ---- 第 6 页：正文开始 ----
c.setFont(FB, 12)
c.drawCentredString(W / 2, y(84), "AGREEMENT OF PURCHASE AND SALE")
base = paragraph(
    "This Agreement of Purchase and Sale is made as of November 5, 2026 by and between Northwind Holdings Inc., a Kentucky "
    "corporation (“Northwind”), and Blue Heron Partners, LLC, a Delaware limited liability company (“Buyer”).",
    112.0, terms=("Northwind", "Buyer"))
c.setFont(FB, 12)
x = 250.5
for ch in "WITNESSETH:":
    c.drawString(x, y(base + 12), ch)
    if ch != ":":
        underline(x, x + stringWidth(ch, FB, 12), base + 12)
    x += stringWidth(ch, FB, 12) + 3.0
base = paragraph(
    "WHEREAS, Northwind and the Asset Selling Corporations are engaged in the Business, and the parties desire that Northwind "
    "sell the Conveyed Assets to Buyer on the terms and subject to the conditions set forth herein;", base + 38.0)
c.setFont(FB, 12)
c.drawCentredString(W / 2, y(base + 14), "ARTICLE 1.")
c.drawCentredString(W / 2, y(base + 14 + PITCH), "DEFINITIONS AND TERMS")
c.drawString(L, y(base + 14 + 2 * PITCH + 12), "Section 1.1.")
c.drawString(L + 72, y(base + 14 + 2 * PITCH + 12), "Definitions.")
base = paragraph("As used in this Agreement, the following terms shall have the meanings set forth or as referenced below:", base + 14 + 3 * PITCH + 24)
definitions = [
    ("Accounts Payable", "“Accounts Payable” shall mean (i) all trade accounts payable and other payment obligations to suppliers of the "
     "Business (excluding intercompany trade accounts and notes payable to Northwind or any of its Affiliates) and the obligation in "
     "respect of all security for such accounts and payment obligations, in each case to the extent relating to the Business."),
    ("Accrued Vacation Compensation", "“Accrued Vacation Compensation” shall mean all Liabilities to be assumed by Buyer with respect to "
     "the accrued vacation of the Transferred Employees determined as of the Closing Date in accordance with the Closing Accounts Principles."),
    ("Closing Date", "“Closing” shall mean the closing of the purchase and sale of the Conveyed Assets, which shall take place on the "
     "third Business Day after the conditions are satisfied (the “Closing Date”), or on such other date as the parties may agree in writing."),
]
for term, text in definitions:
    base = paragraph(text, base + 12.0, terms=(term,))
footer("1")
c.showPage()

# ---- 第 7–8 页：更多定义、定义术语索引 ----
base = 86.0
more = [
    ("Confidentiality Agreement", "“Confidentiality Agreement” shall mean the Confidentiality Agreement between Northwind and Blue Heron "
     "Capital Partners, dated as of July 30, 2026, as amended or supplemented from time to time by the parties thereto."),
    ("Conveyed Assets", "“Conveyed Assets” shall have the meaning set forth in Section 2.2 of this Agreement and shall include the "
     "Assumed Contracts, the Owned Real Property and the Leased Real Property of the Business."),
]
for term, text in more:
    base = paragraph(text, base + 12.0, terms=(term,))
c.setFont(FB, 12)
c.drawString(L, y(base + 24), "Section 1.2.")
c.drawString(L + 72, y(base + 24), "Index of Defined Terms.")
base = paragraph("As used in this Agreement, the following terms shall have the meanings assigned to them in the respective locations "
                 "set forth below:", base + 24 + 26.0)
index = [("Accounts Receivable Facility", "Section 7.1(d)"), ("Acquisition Transaction", "Section 7.15"), ("Actuarial Mediator", "Section 2.11(c)"),
         ("Adjusted Purchase Price", "Section 2.8(d)"), ("Aggregate Cap", "Section 9.6(d)"), ("Alternate Financing", "Section 7.6(c)"),
         ("ASC Transferred Employee", "Section 7.5(c)"), ("Northwind", "Preamble"), ("Northwind Indemnitees", "Section 9.2"),
         ("Asset Allocations", "Section 2.9"), ("Assumed Contracts", "Section 2.2(e)"), ("Assumed Employee Liabilities", "Section 2.5(i)"),
         ("Assumed Liabilities", "Section 2.5"), ("Assumed Tax Liabilities", "Section 2.5(k)"),
         # 第 8 页：整页都是索引（APA 的索引排满三页，这几页被当成正文双栏）
         ("Base Purchase Price", "Section 2.7"), ("Buyer", "Preamble"), ("Buyer Benefit Plans", "Section 7.5(e)"), ("Buyer Fee", "Section 8.2(b)"),
         ("Buyer Indemnitees", "Section 9.1"), ("Closing Accounts", "Section 2.8(b)"), ("Commitment Letters", "Section 6.4(a)"),
         ("Competition Laws", "Section 1.1"), ("Conveyed Real Property", "Section 2.2(a)"), ("Copyrights", "Section 1.1"),
         ("Credit Support Obligations", "Section 7.11(a)"), ("Current Assets", "Section 2.8(a)"), ("Current Liabilities", "Section 2.8(a)"),
         ("Debt Financing", "Section 6.4(a)"), ("Deferred Asset", "Section 2.3(b)"), ("Deferred Liability", "Section 2.3(b)"),
         ("Determination Date", "Section 2.11(b)"), ("Disclosure Schedules", "Article 5"), ("Distributor Agreement", "Section 1.1"),
         ("Effective Time", "Section 1.1"), ("Employee Benefit Plan", "Section 1.1"), ("End Date", "Section 8.1(b)"),
         ("Environmental Laws", "Section 1.1"), ("Estimated Closing Statement", "Section 2.8(a)"), ("European Employees", "Section 7.5(a)"),
         ("Excluded Assets", "Section 2.4"), ("Final Closing Statement", "Section 2.8(c)"), ("Financing Sources", "Section 1.1"),
         ("Fleet Vehicle Leases", "Section 2.2(h)"), ("Foreign Jurisdictions", "Section 7.16")]
c.setFont(F, 12)
head = base + 12.0
c.drawString(L, y(head), "Term")
c.drawString(455, y(head), "Location")
underline(L, L + stringWidth("Term", F, 12), head)
underline(455, 455 + stringWidth("Location", F, 12), head)
base = head + PITCH
for k, (term, loc) in enumerate(index):
    if k == 14:
        footer("2")
        c.showPage()
        base = 86.0
    c.setFont(F, 12)
    c.drawString(L, y(base), term)
    c.drawString(455, y(base), loc)
    base += PITCH
footer("3")
c.save()
