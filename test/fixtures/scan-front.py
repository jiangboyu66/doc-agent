"""生成测试夹具 scan-front.pdf：一页"文字替换式"扫描件的首页（WIPO 专利小册子的首页就是这样排的）。

上半页是左右两栏的著录项目；下面一行通栏的发明名称，词距很宽，恰好有一个词距落在两栏的栏间距上；再一行通栏的名称；
下半页左边是附图（只在扫描图里，右缘略微越过栏间距的中线），右边是两端对齐的摘要；左边距里竖排一条条码。

易错点（WO2017038541A1 首页，2026-11-27）：
- 附图按栏间距的中线算成"通栏"：右栏的摘要从附图处切成上下两段、各成一节，开头几行排在附图上面；
- 左栏只有附图、右栏是摘要时，可编辑排版让 Word 平衡两栏，摘要的开头几行排到了左栏附图的下面。
（原件里发明名称那一行还被切成了左右两栏的两行；这里的发明名称在转换时合成了跨过栏间距的文字片段，不复现那一处，
那一处以原件的转换结果核对。）"""
import io
import sys

from PIL import Image, ImageDraw
from pypdf import PdfReader, PdfWriter
from pypdf.generic import DecodedStreamObject, NameObject, NumberObject
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas

W, H = 595.0, 842.0
K = 150 / 72  # 图上每点的像素数
PW, PH = round(W * K), round(H * K)
out = sys.argv[1] if len(sys.argv) > 1 else "scan-front.pdf"
draft = io.BytesIO()
c = canvas.Canvas(draft, pagesize=(W, H), pageCompression=1, invariant=1)
c.setCreator("doc-agent test fixture")

im = Image.new("L", (PW, PH), 255)
dr = ImageDraw.Draw(im)


def box(x0, y0, x1, y1, width=2):
    dr.rectangle([x0 * K, y0 * K, x1 * K, y1 * K], outline=0, width=width)


# 附图：一个大框里几个小框，用线连起来；右缘 322，略微越过栏间距的中线（约 315），没碰到右栏的字（335 起）
box(84, 420, 322, 570)
for k in range(3):
    box(100 + k * 72, 450, 150 + k * 72, 500)
    if k:
        dr.line([(150 + (k - 1) * 72) * K, 475 * K, (100 + k * 72) * K, 475 * K], fill=0, width=2)
box(120, 520, 290, 555)
# 左边距里竖排的条码
for k in range(80):
    w = 1 + (k * 5) % 3
    y = 600 + k * 2.5
    dr.rectangle([22 * K, y * K, 44 * K, (y + w * 0.6) * K], fill=0)

buf = io.BytesIO()
im.save(buf, "PNG")
buf.seek(0)
c.drawImage(ImageReader(buf), 0, 0, W, H)


def text(x, y, s, size=9, font="Helvetica"):
    c.setFont(font, size)
    c.drawString(x, H - y, s)


def spread(x0, x1, y, s, size=9):
    """两端对齐的一行：字距拉开，使这一行从 x0 排到 x1"""
    natural = pdfmetrics.stringWidth(s, "Helvetica", size)
    t = c.beginText(x0, H - y)
    t.setFont("Helvetica", size)
    t.setCharSpace((x1 - x0 - natural) / max(1, len(s) - 1))
    t.textOut(s)
    c.drawText(t)


# 著录项目：左栏从 74 排到 305（两端对齐的识别文字），右栏从 335 起
LEFT = ["(51) International Patent Classification:", "H04N 5/232 (2006.01)  H04N 5/225 (2006.01)", "(21) International Application Number:",
        "PCT/ZZ2099/000001", "(22) International Filing Date: 23 August 2099", "(25) Filing Language: English", "(26) Publication Language: English",
        "(30) Priority Data: 2098-174018  3 September 2098", "(71) Applicant: EXAMPLE CORPORATION", "(72) Inventors: A. EXAMPLE; B. SAMPLE"]
RIGHT = ["(74) Agent: C. AGENT et al.", "(81) Designated States (unless otherwise indicated,", "for every kind of national protection",
         "available): AE, AG, AL, AM, AO, AT, AU, AZ, BA,", "BB, BG, BH, BN, BR, BW, BY, BZ, CA, CH, CL,", "(84) Designated States (unless otherwise indicated,",
         "for every kind of regional protection available):", "ARIPO (BW, GH, GM, KE, LR, LS, MW, MZ, NA,", "RW, SD, SL, ST, SZ, TZ, UG, ZM, ZW), Eurasian",
         "(AM, AZ, BY, KG, KZ, RU, TJ, TM), European"]
for k, s in enumerate(LEFT):
    spread(74, 305, 140 + k * 16, s)
for k, s in enumerate(RIGHT):
    spread(335, 550, 143 + k * 16, s)
text(470, 312, "[Continued on next page]")
# 发明名称：通栏一行，识别文字的词距很宽；"PROCESSING" 到 316 为止，"METHOD," 从 330 起：恰好有一个词距落在栏间距里
for x, word in ((74, "(54)"), (92, "Title:"), (118, "VIDEO"), (142, "PROCESSING"), (196, "DEVICE,"), (228, "VIDEO"), (262, "PROCESSING"),
                (330, "METHOD,"), (366, "AND"), (388, "PROGRAM")):
    text(x, 336, word, 7)
text(74, 352, "(54) Name of the invention: video processing device, video processing method and program", 9)
# 附图里的标注是识别出来的文字，画在扫描图上面（WIPO 的做法）
text(88, 432, "10 Video processing device", 6)
for k, label in enumerate(("Input unit", "Tracking", "Output unit")):
    text(104 + k * 72, 478, label, 6)
text(160, 540, "Control unit", 6)
# 摘要：右栏，两端对齐（从 335 排到 550）
ABSTRACT = ["(57) Abstract: The present technology pertains to a video", "processing device, a video processing method and a", "program with which it is possible to provide a video",
            "centred on a subject in cases when the subject moves.", "This video processing device acquires viewpoint", "position information indicating the movement range",
            "of a viewpoint position relative to a subject of", "interest included in an input image frame and time", "information indicating the temporal range of the",
            "time sequence with which the input image frame is", "filmed, and processes the input image frame so that", "the subject of interest is located at a specific",
            "position in an output image frame. The present", "technology can be applied to a video processing"]
for k, s in enumerate(ABSTRACT):
    spread(335, 550, 400 + k * 11, s)
text(335, 400 + len(ABSTRACT) * 11, "device for processing a bullet-time video.")
c.showPage()
c.save()

# reportlab 只会写 8 位灰度图；真正的扫描件是 1 位黑白图，逐张改写
writer = PdfWriter(clone_from=PdfReader(io.BytesIO(draft.getvalue())))
for pg in writer.pages:
    images = pg["/Resources"]["/XObject"]
    for name in list(images):
        old = images[name].get_object()
        if old.get("/Subtype") != "/Image":
            continue
        w, h = int(old["/Width"]), int(old["/Height"])
        bits = Image.frombytes("L", (w, h), old.get_data()).point(lambda v: 255 if v >= 128 else 0).convert("1")
        new = DecodedStreamObject()
        new.set_data(bits.tobytes())
        for key, value in (("/Type", NameObject("/XObject")), ("/Subtype", NameObject("/Image")), ("/Width", NumberObject(w)), ("/Height", NumberObject(h)),
                           ("/ColorSpace", NameObject("/DeviceGray")), ("/BitsPerComponent", NumberObject(1))):
            new[NameObject(key)] = value
        images[NameObject(name)] = writer._add_object(new.flate_encode())
with open(out, "wb") as f:
    writer.write(f)
