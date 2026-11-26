"""生成测试夹具 scan-text.pdf：6 页"文字替换式"扫描件（WIPO 公开的专利小册子就是这种做法）。
每页一张整页的黑白扫描图；识别出来的字从图上擦掉，改成真正的文字画在图上面，图里只剩没换成文字的东西：
页眉（公开号）、附图、个别没认出来的字。图和文字互补，哪一层单独看都不完整。

第 1 页  正文 + 一幅框图 + 句子中间一个只在图里的字
第 2 页  正文 + 两个只在图里的字（权利要求的编号）
第 3 页  正文
第 4 页  只有附图，没有文字层
第 5 页  没做识别的整页文字扫描（没有文字层，但看得出是文字）
第 6 页  带表格线的表单，格子里是真正的文字

易错点：把整页图衬在文字后面，Word 里文字一重排就和图错位；把图直接丢掉，附图和那几个字就没了。
应当把图拆开：附图单独成图，句中的字留在原来的位置，每页重复的页眉去掉，没有文字层的页按整页图保留并如实上报。"""
import io
import sys

from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfReader, PdfWriter
from pypdf.generic import DecodedStreamObject, NameObject, NumberObject
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas

W, H = 595.0, 842.0
K = 150 / 72  # 图上每点的像素数
PW, PH = round(W * K), round(H * K)
SANS = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
out = sys.argv[1] if len(sys.argv) > 1 else "scan-text.pdf"
draft = io.BytesIO()
c = canvas.Canvas(draft, pagesize=(W, H), pageCompression=1)
c.setCreator("doc-agent test fixture")

BODY = [
    "The image processing device receives the input frames from the cameras and stores them",
    "in the frame memory together with the capture time of each frame so that later stages",
    "can select the frames that belong to the same instant and correct their viewing angle",
    "before the frames are combined into one output stream for the display or the recorder",
]


def page_image():
    im = Image.new("L", (PW, PH), 255)
    return im, ImageDraw.Draw(im)


def ink(dr, x, y, text, size):
    """把字画进扫描图：x 是左边，y 是基线（PDF 点，自上而下）。"""
    font = ImageFont.truetype(SANS, round(size * K))
    dr.text((x * K, y * K), text, fill=0, font=font, anchor="ls")


def header(dr, page):
    ink(dr, 60, 40, "WO 2099/000001", 10)
    ink(dr, W / 2 - 4, 40, str(page), 10)
    ink(dr, W - 60 - 96, 40, "PCT/ZZ2099/000001", 10)


TEXT = []  # 这一页的文字：(x, 基线, 字号, 内容)，等扫描图铺好以后再画在它上面


def place(im):
    buf = io.BytesIO()
    im.save(buf, "PNG")
    buf.seek(0)
    c.drawImage(ImageReader(buf), 0, 0, W, H)
    for x, y, size, text in TEXT:
        c.setFont("Helvetica", size)
        c.drawString(x, H - y, text)
    TEXT.clear()
    c.showPage()


PARAGRAPH = 4  # 每段的行数：末行短，以句号结束


def lines(dr, y, count, start=0, gaps=None):
    """画 count 行正文。gaps[k] = (前半句, 图里的字, 后半句)：这一行中间那个字只在扫描图里。"""
    for k in range(count):
        gap = (gaps or {}).get(k)
        last = k % PARAGRAPH == PARAGRAPH - 1
        if gap:
            before, glyph, after = gap
            TEXT.append((60, y, 11, before))
            x = 60 + pdfmetrics.stringWidth(before, "Helvetica", 11) + 1.5
            ink(dr, x, y, glyph, 11)
            TEXT.append((x + pdfmetrics.stringWidth(glyph, "Helvetica", 11) + 2.5, y, 11, after))
        else:
            text = BODY[(start + k) % len(BODY)]
            TEXT.append((60, y, 11, text[:52].rstrip() + "." if last else text))
        y += 16 + (10 if last else 0)
    return y


def diagram(dr, x, y, w, h, boxes=3):
    """框图：几个方框用线连起来，框里的字也是图。"""
    step = w / boxes
    for k in range(boxes):
        x0, x1 = x + k * step + 10, x + (k + 1) * step - 10
        dr.rectangle([x0 * K, (y + 20) * K, x1 * K, (y + h - 20) * K], outline=0, width=3)
        ink(dr, x0 + 12, y + h / 2 + 4, f"S{k + 1}", 14)
        if k:
            dr.line([(x0 - 20) * K, (y + h / 2) * K, x0 * K, (y + h / 2) * K], fill=0, width=3)


# 第 1 页
im, dr = page_image()
header(dr, 1)
y = lines(dr, 80, 8)
diagram(dr, 110, y + 6, 360, 150)
y = lines(dr, y + 180, 8, 1, {2: ("is published on 9 March 2099 as document number ", "7", " of the series and")})
lines(dr, y, 12, 2)
place(im)

# 第 2 页
im, dr = page_image()
header(dr, 2)
y = lines(dr, 80, 12, 1, {2: ("The device according to claim ", "1", ", wherein the frames are stored"), 6: ("The device according to claim ", "2", ", wherein the angle is corrected")})
lines(dr, y, 24, 3)
place(im)

# 第 3 页
im, dr = page_image()
header(dr, 3)
lines(dr, 80, 36, 2)
place(im)

# 第 4 页：只有附图
im, dr = page_image()
header(dr, 4)
ink(dr, W / 2 - 20, 110, "FIG. 1", 14)
diagram(dr, 70, 150, 450, 220, 3)
diagram(dr, 70, 420, 450, 220, 2)
place(im)

# 第 5 页：没做识别的文字扫描
im, dr = page_image()
header(dr, 5)
y = 80
for k in range(40):
    ink(dr, 60, y, BODY[k % len(BODY)], 11)
    y += 16
place(im)

# 第 6 页：表单
im, dr = page_image()
header(dr, 6)
rows, x0, x1, top = 12, 60, W - 60, 80
for r in range(rows + 1):
    dr.line([x0 * K, (top + r * 50) * K, x1 * K, (top + r * 50) * K], fill=0, width=3)
for x in (x0, 200, x1):
    dr.line([x * K, top * K, x * K, (top + rows * 50) * K], fill=0, width=3)
for r in range(rows):
    TEXT.append((x0 + 6, top + r * 50 + 20, 10, f"Box No. {r + 1}"))
    TEXT.append((206, top + r * 50 + 20, 10, "The applicant confirms the statement made in this box of the form"))
    TEXT.append((206, top + r * 50 + 34, 10, "and attaches the documents that are listed in the annex to it"))
place(im)
c.save()

# reportlab 只会写 8 位灰度图；真正的扫描件是 1 位黑白图，逐张改写
writer = PdfWriter(clone_from=PdfReader(io.BytesIO(draft.getvalue())))
for page in writer.pages:
    images = page["/Resources"]["/XObject"]
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
