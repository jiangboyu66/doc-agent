"""生成测试夹具 scan-strip.pdf：3 页"文字替换式"扫描件，只有第 2 页页顶有一条只在扫描图里的公开号条（文字 + 条码）。

WIPO 专利小册子的节选就是这样：首页顶上没有公开号条，第 2 页起才有；只节选几页时，这条只出现在一两页上，
不算每页重复的页眉，留在正文里。可编辑排版不保留原文的分页，第 1 页排满（没有提前换页）时，第 2 页的内容接着排，
页顶的公开号条就跑到了第 1 页正文的末尾（WO2017038541A1 节选，2026-11-27）。应当让第 2 页从新的一页开始。

第 1 页  首页：顶上一行标题（比其他页的正文高），正文排满整页，最后一行排到右边界（段落接到下一页）
第 2 页  页顶一条公开号条（"WO 2099/000002" 和一段条码，都只在扫描图里），落在首页标题以下，即正文区域里；下面是正文
第 3 页  正文，页顶没有公开号条"""
import io
import sys

from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfReader, PdfWriter
from pypdf.generic import DecodedStreamObject, NameObject, NumberObject
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

W, H = 595.0, 842.0
K = 150 / 72  # 图上每点的像素数
PW, PH = round(W * K), round(H * K)
SANS = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
out = sys.argv[1] if len(sys.argv) > 1 else "scan-strip.pdf"
draft = io.BytesIO()
c = canvas.Canvas(draft, pagesize=(W, H), pageCompression=1, invariant=1)
c.setCreator("doc-agent test fixture")

BODY = [
    "The image processing device receives the input frames from the cameras and stores them",
    "in the frame memory together with the capture time of each frame so that later stages",
    "can select the frames that belong to the same instant and correct their viewing angle",
    "before the frames are combined into one output stream for the display or the recorder",
]


def page(strip, first_line, count, title=None):
    im = Image.new("L", (PW, PH), 255)
    dr = ImageDraw.Draw(im)
    if strip:
        # 公开号条：文字与条码都只在扫描图里
        font = ImageFont.truetype(SANS, round(12 * K))
        dr.text((60 * K, 66 * K), "WO 2099/000002", fill=0, font=font, anchor="ls")
        x = 200
        for k in range(60):
            w = 1 + (k * 7) % 3
            dr.rectangle([x * K, 52 * K, (x + w * 0.6) * K, 68 * K], fill=0)
            x += w * 0.6 + 1.2
    buf = io.BytesIO()
    im.save(buf, "PNG")
    buf.seek(0)
    c.drawImage(ImageReader(buf), 0, 0, W, H)
    if title:
        c.setFont("Helvetica-Bold", 12)
        c.drawString(60, H - 40, title)
    c.setFont("Helvetica", 11)
    y = 90
    for k in range(count):
        c.drawString(60, H - y, BODY[(first_line + k) % len(BODY)])
        y += 16
    c.showPage()


# 44 行：从 90 排到 778，排满整页；每行都排到右边界（没有段落在页尾结束）
page(False, 0, 44, "(12) INTERNATIONAL APPLICATION PUBLISHED UNDER THE TREATY")
page(True, 1, 44)
page(False, 2, 44)
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
