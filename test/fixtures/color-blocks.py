"""生成测试夹具 color-blocks.pdf：带底色的文字块。两页都铺着浅米色的页面底色；
第 1 页：深色横幅（白色标题 + 副标题）、一行三个彩色指标卡片（白字）、左边带色条的提示框、正文；
第 2 页：左侧深色侧边栏（白字的联系方式、技能、两行一组的条目）+ 右侧正文。
以前：可编辑排版不画正文区域的色块，白字落在白纸上看不见；侧边栏这样带文字的色块被当成图整块截成图片，文字不能编辑；
页面底色接近白色，没有转换。"""
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color

out = sys.argv[1] if len(sys.argv) > 1 else "color-blocks.pdf"
W, H = 595, 842
c = canvas.Canvas(out, pagesize=(W, H))
cream = Color(1, 0.992, 0.965)
navy = Color(0.06, 0.17, 0.30)
white = Color(1, 1, 1)
body = [
    "Cities are rethinking the street. Where cars once dominated, planners now carve out room for buses,",
    "bikes and people on foot, and the results show up in surprising places: corner shops report steadier",
    "trade, emergency crews reach calls faster, and summer heat on the pavement drops by several degrees",
    "once trees return. None of this happened by accident; it took years of patient work by many people.",
]


def page_bg():
    c.setFillColor(cream)
    c.rect(0, 0, W, H, stroke=0, fill=1)


def y(v):
    return H - v


# 第 1 页
page_bg()
c.setFillColor(navy)
c.rect(57, y(130), 481, 72, stroke=0, fill=1)
c.setFillColor(white)
c.setFont("Helvetica-Bold", 20)
c.drawString(72, y(88), "Quarterly Operations Review")
c.setFont("Helvetica", 9)
c.drawString(72, y(110), "Internal - prepared for the executive committee - Q1 2026")
cards = [((0.18, 0.53, 0.67), "98.7%", "Service uptime"), ((0.95, 0.56, 0.0), "4.2 h", "Median ticket resolution"), ((0.23, 0.55, 0.35), "-11%", "Unit cost vs. Q4")]
for k, (col, big, small) in enumerate(cards):
    x = 63 + k * 160
    c.setFillColor(Color(*col))
    c.rect(x, y(207), 150, 66, stroke=0, fill=1)
    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(x + 8, y(166), big)
    c.setFont("Helvetica", 8)
    c.drawString(x + 8, y(190), small)
c.setFillColor(Color(1, 0.94, 0.86))
c.rect(57, y(262), 481, 40, stroke=0, fill=1)
c.setFillColor(Color(0.95, 0.56, 0.0))
c.rect(57, y(262), 4, 40, stroke=0, fill=1)
c.setFillColor(Color(0, 0, 0))
c.setFont("Helvetica-Bold", 9)
c.drawString(70, y(238), "Decision needed:")
c.setFont("Helvetica", 9)
c.drawString(150, y(238), "approve the migration of the billing cluster before May 15")
c.drawString(70, y(252), "to avoid the renewal of the legacy contract.")
c.setFont("Helvetica", 10)
for i, s in enumerate(body * 3):
    c.drawString(57, y(290 + 14 * i), s)
c.showPage()

# 第 2 页：侧边栏
page_bg()
c.setFillColor(navy)
c.rect(0, 0, 190, H, stroke=0, fill=1)
c.setFillColor(white)
side = [("Helvetica-Bold", 11, "CONTACT"), ("Helvetica", 8.5, "Phone 138-0000-0000"), ("Helvetica", 8.5, "zhang.wei@example.com"),
        ("Helvetica", 8.5, "Pudong, Shanghai"), ("Helvetica-Bold", 11, "SKILLS"), ("Helvetica", 8.5, "Python / SQL"),
        ("Helvetica", 8.5, "Spark / Flink"), ("Helvetica", 8.5, "Kafka"), ("Helvetica-Bold", 11, "EDUCATION"),
        ("Helvetica", 8.5, "MSc Computer Science"), ("Helvetica", 8.5, "Example University, 2014 - 2017")]
for i, (f, sz, t) in enumerate(side):
    c.setFont(f, sz)
    c.drawString(24, y(160 + 24 * i), t)
c.setFillColor(Color(0, 0, 0))
c.setFont("Helvetica-Bold", 20)
c.drawString(220, y(90), "Zhang Wei")
c.setFont("Helvetica", 10)
for i, s in enumerate(["Senior data engineer with eight years of experience in batch and",
                       "stream processing, data governance and the design of large pipelines.",
                       "Led a team of six engineers and migrated two hundred jobs to the cloud.",
                       "Speaks at meetups about data quality and cost control in data platforms."]):
    c.drawString(220, y(130 + 15 * i), s)
c.showPage()
c.save()
