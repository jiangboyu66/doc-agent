"""生成测试夹具 scan-front-words.pdf：WIPO 小册子首页那样的"文字替换式"扫描件，识别文字按词逐个画在扫描图上面。

与 scan-front.pdf 的区别：识别文字层是一个词一个文字片段，按原件的位置摆放（不是整行一个片段、用字距拉开）。
窄栏两端对齐的行，词距被拉开到 5–14pt，常有一两处超过 12pt；这样的空隙不是制表间距。

易错点（WO2017038541A1 首页，2026-11-27）：
- 摘要、著录项目 (72)(74) 里超过 12pt 的词距被当成制表位：这一行单独成段、在原位置设制表位，上下行也并不进来，
  译文变短后行中、行尾留下大片空白，一句话被拆到几段里分别翻；
- 词距均匀但偏宽的英文发明名称一行，同样被拆成几段制表位；
- 右栏以 "(74)" 开头，左右两栏之间没有分栏符：右栏从哪一行开始由 Word / LibreOffice 平衡两栏决定，
  排到栏顶的 (74) 段前距被去掉，比左栏的 (51) 高出一行。
真正的对齐字段（IPC 两列、申请号靠右）空隙在 28pt 以上，仍是制表位。"""
import io
import sys

from PIL import Image, ImageDraw
from pypdf import PdfReader, PdfWriter
from pypdf.generic import DecodedStreamObject, NameObject, NumberObject
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas

W, H = 595.0, 842.0
K = 150 / 72
PW, PH = round(W * K), round(H * K)
out = sys.argv[1] if len(sys.argv) > 1 else "scan-front-words.pdf"
draft = io.BytesIO()
c = canvas.Canvas(draft, pagesize=(W, H), pageCompression=1, invariant=1)
c.setCreator("doc-agent test fixture")

im = Image.new("L", (PW, PH), 255)
dr = ImageDraw.Draw(im)


def box(x0, y0, x1, y1, width=2):
    dr.rectangle([x0 * K, y0 * K, x1 * K, y1 * K], outline=0, width=width)


# 页眉下的横线、附图（左栏下半）
dr.line([74 * K, 142 * K, 552 * K, 142 * K], fill=0, width=2)
box(84, 380, 300, 520)
for k in range(3):
    box(100 + k * 66, 410, 146 + k * 66, 450)
box(120, 470, 270, 500)
buf = io.BytesIO()
im.save(buf, "PNG")
buf.seek(0)
c.drawImage(ImageReader(buf), 0, 0, W, H)


def word(x, y, s, size):
    c.setFont("Helvetica", size)
    c.drawString(x, H - y, s)
    return x + pdfmetrics.stringWidth(s, "Helvetica", size)


def line(x0, x1, y, words, size, extra=()):
    """一行词，从 x0 排到 x1（两端对齐）；extra：(第几个词距, 多加的宽度)，模拟识别文字层词宽不准造成的参差"""
    widths = [pdfmetrics.stringWidth(w, "Helvetica", size) for w in words]
    gaps = len(words) - 1
    add = dict(extra)
    base = (x1 - x0 - sum(widths) - sum(add.values())) / max(1, gaps)
    x = x0
    for k, w in enumerate(words):
        x = word(x, y, w, size)
        if k < gaps:
            x += base + add.get(k, 0)


def natural(x, y, words, size, extra=()):
    """不拉开的一行（段落末行）；extra 同上"""
    add = dict(extra)
    space = pdfmetrics.stringWidth(" ", "Helvetica", size)
    for k, w in enumerate(words):
        x = word(x, y, w, size) + space + add.get(k, 0)


# 左栏：著录项目
word(74, 156, "(51)", 6.5)
word(95, 156, "International Patent Classification:", 8)
for x, s in ((95, "H04N"), (116, "5/232"), (138, "(2006.01)"), (199, "H04N"), (220, "5/225"), (242, "(2006.01)")):
    word(x, 167, s, 6.5)
word(74, 192, "(21)", 6.5)
word(95, 192, "International Application Number:", 8)
word(305 - pdfmetrics.stringWidth("PCT/ZZ2099/000001", "Helvetica", 6.5), 192, "PCT/ZZ2099/000001", 6.5)
# (72)：两端对齐的三行，词距被拉开，有几处超过 12pt
word(74, 220, "(72)", 6.5)
line(95, 305, 220, "Inventors: EXAMPLE Taro; 1-7-1 Konan, Minato-ku,".split(), 8, extra=((3, 4),))
line(95, 305, 230.5, "Tokyo 108-0075 (ZZ). SAMPLE Jiro; 1-7-1".split(), 8, extra=((2, 3),))
natural(95, 241, "Konan, Minato-ku, Tokyo (ZZ).".split(), 8, extra=((0, 7), (1, 12), (2, 7)))
# 右栏：(74) 与 (51) 同高；第一行两端对齐（识别文字的词距 10pt 上下），有一处 18pt 的空隙
word(320, 155, "(74)", 6.5)
line(339, 550, 155, "Agent: AGENT Saburo et al.; Nishi-Shinjuku,".split(), 8, extra=((3, 8),))
natural(339, 165.5, "7-5-25 Shinjuku-ku, Tokyo 160-0023 (ZZ).".split(), 8)
word(320, 190, "(81)", 6.5)
line(339, 550, 190, "Designated States (unless otherwise indicated, for every".split(), 8)
line(339, 550, 200.5, "kind of national protection available): AE, AG, AL, AM,".split(), 8)
natural(339, 211, "AO, AT, AU, AZ.".split(), 8)
# 通栏的英文发明名称：词距均匀而偏宽（10–14pt）
x = 74
for s in ("(54)", "Title:", "VIDEO", "PROCESSING", "DEVICE,", "VIDEO", "PROCESSING", "METHOD,", "AND", "PROGRAM"):
    x = word(x, 330, s, 6.5) + (5.6 if s == "(54)" else 11 + (len(s) % 3))
word(74, 345, "(54) Name of the invention: video processing device, video processing method and program", 8)
# 右栏的摘要：两端对齐，词距 5–14pt，十行里五行有一处超过 12pt
ABSTRACT = ["(57) Abstract: The present technology pertains to a video", "processing device, a video processing method, and a",
            "program with which it is possible to provide a bullet-time", "video centering on a subject in cases when the subject",
            "moves. This video processing device acquires viewpoint", "position information indicating the movement range of a",
            "viewpoint position relative to a subject of interest", "included in an input image frame, and time information",
            "indicating the temporal range of the time sequence with", "which the input image frame is filmed, and processes the"]
for k, s in enumerate(ABSTRACT):
    words = s.split()
    line(335, 550, 385 + k * 10.2, words, 6.5, extra=((len(words) // 2, 5),) if k % 2 else ())
natural(335, 385 + len(ABSTRACT) * 10.2, "input image frame for a bullet-time video.".split(), 6.5)
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
