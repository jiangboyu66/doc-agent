"""生成测试夹具 definitions.pdf：港股招股书摘录（虚构公司），两页。
  第 1 页：封面的"项目 : 内容"清单（项目名称折行、内容多行，冒号单独一列），下面是一段通栏说明；
  第 2 页："释义"清单：左栏词条（有的两行），右栏释义（有的多行），词条与释义第一行同高；
          一行的条目之间有段距，也有两条之间没有段距、只靠引号分开的。
  第 3 页：真正的正文双栏（两栏等宽、段落各排各的），仍应用 Word 分栏还原。
  第 4 页：邮件抬头（From / To / Subject，内容从版心 15% 处起排），不会被分栏，仍写成按制表位对齐的段落。
易错点：1.1.2 把封面清单和释义清单当成正文双栏，Word 里词条连续排在左栏、释义连续排在右栏，
        读者对不上哪条释义属于哪个词条，相邻的释义还会并成一段（2026-11-27 LTB-07）。
标准答案：清单写成无边框表格，一条一行（封面清单三列：项目、冒号、内容；释义两列）；第 3 页仍是分栏；
        第 4 页的邮件抬头不改成表格（1.1.2 就是对的）。"""
import sys
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

W, H = 595.0, 842.0
out = sys.argv[1] if len(sys.argv) > 1 else "definitions.pdf"
c = canvas.Canvas(out, pagesize=(W, H))
F, FB = "Times-Roman", "Times-Bold"


def top(y: float) -> float:
    """以页面顶端为原点的纵坐标 → PDF 坐标（基线）。"""
    return H - y


# ---- 第 1 页：封面清单 ----
c.setFont(FB, 17)
c.drawCentredString(W / 2, top(170), "Northwind Polar Logistics Holdings Limited")
c.setFont(FB, 14)
c.drawCentredString(W / 2, top(235), "GLOBAL OFFERING")
c.setFont(F, 10)
cover = [
    (["Number of Offer Shares under the Global", "Offering"], ["80,000,000 Shares (subject to the Over-allotment", "Option)"]),
    (["Number of Hong Kong Offer Shares"], ["8,000,000 Shares (subject to reallocation)"]),
    (["Number of International Offer Shares"], ["72,000,000 Shares (subject to reallocation and the", "Over-allotment Option)"]),
    (["Maximum Offer Price"], ["HK$3.20 per Offer Share, plus brokerage of 1.0%, SFC", "transaction levy of 0.0027% and Stock Exchange trading", "fee of 0.00565%"]),
    (["Nominal value"], ["US$0.0001 per Share"]),
    (["Stock code"], ["[TBC]"]),
]
y = 262.0
for label, value in cover:
    c.drawString(277.5, top(y), ":")
    for k in range(max(len(label), len(value))):
        if k < len(label):
            c.drawString(70.9, top(y + 13.4 * k), label[k])
        if k < len(value):
            c.drawString(283.5, top(y + 13.4 * k), value[k])
    y += 13.4 * max(len(label), len(value)) + 3
c.setFont(F, 8)
para = ("Hong Kong Exchanges and Clearing Limited, The Stock Exchange of Hong Kong Limited and Hong Kong Securities "
        "Clearing Company Limited take no responsibility for the contents of this synthetic prospectus and make no "
        "representation as to its accuracy or completeness.")
words, line, yy = para.split(), "", y + 30
for w_ in words:
    if stringWidth(line + " " + w_, F, 8) > 453:
        c.drawString(70.9, top(yy), line)
        yy += 10.6
        line = w_
    else:
        line = (line + " " + w_).strip()
c.drawString(70.9, top(yy), line)
c.showPage()

# ---- 第 2 页：释义 ----
c.setFont(FB, 12)
c.drawCentredString(W / 2, top(50), "DEFINITIONS")
c.setFont(F, 10)
c.drawString(70.9, top(87), "In this prospectus, unless the context otherwise requires, the following expressions have the following")
c.drawString(70.9, top(99.4), "meanings.")
c.setFont(F, 9)
defs = [
    (["“Accountants’ Report”"], ["the accountants’ report of our Group set out in Appendix I to this prospectus"]),
    (["“Articles”"], ["the amended and restated articles of association of our Company conditionally", "adopted on 3 November 2026 and effective on the Listing Date, as amended from", "time to time"]),
    (["“Board”"], ["the board of Directors"]),
    (["“Business Day”"], ["a day on which banks in Hong Kong are generally open for normal banking business", "to the public"]),
    (["“Companies (Winding Up and", "Miscellaneous Provisions) Ordinance”"], ["the Companies (Winding Up and Miscellaneous Provisions) Ordinance (Chapter 32", "of the Laws of Hong Kong), as amended, supplemented or otherwise modified from", "time to time"]),
    (["“Director(s)”"], ["the director(s) of our Company"]),
    (["“Group”"], ["our Company and its subsidiaries"]),
    (["“HK$”"], ["Hong Kong dollars, the lawful currency of Hong Kong"]),
    (["“Listing”"], ["the listing of the Shares on the Main Board"]),
    (["“PRC”"], ["the People’s Republic of China, excluding, for the purpose of this prospectus,", "Hong Kong, Macau and Taiwan"]),
]
y = 123.0
for term, meaning in defs:
    for k in range(max(len(term), len(meaning))):
        if k < len(term):
            c.drawString(70.9, top(y + 12.3 * k), term[k])
        if k < len(meaning):
            c.drawString(224.6, top(y + 12.3 * k), meaning[k])
    y += 12.3 * max(len(term), len(meaning)) + 3.4
# 两条之间没有段距：只靠引号分开
for term, meaning in [("“SFC”", "the Securities and Futures Commission of Hong Kong"),
                      ("“SFO”", "the Securities and Futures Ordinance (Chapter 571 of the Laws of Hong Kong)"),
                      ("“US$”", "United States dollars, the lawful currency of the United States")]:
    c.drawString(70.9, top(y), term)
    c.drawString(224.6, top(y), meaning)
    y += 12.3
c.showPage()

# ---- 第 3 页：正文双栏 ----
c.setFont(FB, 12)
c.drawCentredString(W / 2, top(50), "INDUSTRY OVERVIEW")
c.setFont(F, 9)
body = ("The cold chain logistics market in the synthetic region grew steadily over the track record period, driven by "
        "rising demand for fresh food, pharmaceuticals and other temperature sensitive goods. Operators compete on network "
        "coverage, temperature control and digital capability. Larger operators have invested in automated warehouses and "
        "refrigerated fleets, while smaller operators focus on regional routes. Industry participants expect further "
        "consolidation as customers favour providers able to offer integrated warehousing and transportation services. ")
cols = [(70.9, 287.0), (308.0, 524.4)]
for ci, (x0, x1) in enumerate(cols):
    y = 80.0 + 6 * ci
    for p in range(4):
        words, line = (body * 2).split(), ""
        n = 0
        for w_ in words:
            if stringWidth(line + " " + w_, F, 9) > x1 - x0:
                c.drawString(x0, top(y), line)
                y += 11.5
                n += 1
                line = w_
                if n >= 9 + p + ci:
                    break
            else:
                line = (line + " " + w_).strip()
        y += 8
c.showPage()

# ---- 第 4 页：邮件抬头（内容从版心 15% 处起排，不会被分栏），仍是按制表位对齐的段落 ----
c.setFont(F, 10)
header = [
    ("From:", ["Alex Chan <alex.chan@northwind.example>"]),
    ("Sent:", ["Wednesday, 4 November 2026 9:12 AM"]),
    ("To:", ["Sam Lee <sam.lee@harbourlaw.example>"]),
    ("Cc:", ["Jo Wong <jo.wong@northwind.example>; Kim Ho <kim.ho@harbourlaw.example>;", "legal@northwind.example"]),
    ("Subject:", ["RE: Synthetic Holdings Limited v Northwind Polar Logistics Holdings Limited", "[WITHOUT PREJUDICE SAVE AS TO COSTS]"]),
    ("Attachments:", ["Draft settlement agreement (v3).docx"]),
]
y = 82.0
for label, value in header:
    c.drawString(57.0, top(y), label)
    for k, v in enumerate(value):
        c.drawString(129.4, top(y + 11.3 * k), v)
    y += 11.3 * len(value) + 2.2
c.drawString(57.0, top(y + 24), "Dear Sam,")
c.drawString(57.0, top(y + 46), "We refer to your email of 3 November 2026 and set out our client's position below.")
c.showPage()
c.save()
