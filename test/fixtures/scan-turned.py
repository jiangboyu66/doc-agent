"""生成测试夹具 scan-turned.pdf：没有文字层的扫描专利（US20040163766A1 的 incoPat 文本就是这样的，2026-11-27）。

每页一张 150 dpi 的灰度扫描图，没有文字层：
  第 1 页  纵向页面上的文字页：页眉一行、页码；左右两栏正文，段首是 [0001] 这样的编号，编号后空得很宽；左栏中间一行居中的
          全大写小标题（字母挨着）；右栏下面一幅附图（外框里有一行标注，框外几个附图标记连着引线）。
  第 2 页  横向页面上的附图页：附图扫成竖幅位图，再转 90° 画到横向页面上（cm 0 -h w 0 0 h）；页边一条转了 90° 的页眉。
  第 3 页  页面框与扫描图横竖互换：纵向的页面框里转 90° 画了一张横幅的附图，图比页面宽，阅读器里右边被裁掉一截。
  第 4 页  纵向页面上的表格页：粗表格线，格里是文字和数值；表下"FIG. 4"。

易错点（转换的引擎与 AltaLex 翻译，2026-11-27）：
- 转了 90° 放的附图页拆不开（只拆正放的扫描图），被当成"没读出来的文字页"，识别后整页图换成一列附图标记——附图没了；
- 页面框横竖互换的页，图写进 Word 时没转正、比页面宽，横躺着被裁掉；
- 没有文字层的文字页整页图换成识别出的文字后重排，两栏变一栏、附图丢了。
"""
import io
import sys

from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfReader, PdfWriter
from pypdf.generic import DecodedStreamObject, NameObject, NumberObject
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

DPI = 150
K = DPI / 72  # 图上每点的像素数
SERIF = "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf"
SERIF_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"
SANS = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
out = sys.argv[1] if len(sys.argv) > 1 else "scan-turned.pdf"


def font(path, pt):
    return ImageFont.truetype(path, round(pt * K))


def bw(im):
    """纯黑白（扫描件的 1 位图，存成只有 0 / 255 的灰度）"""
    return im.point(lambda v: 0 if v < 160 else 255)


def ink(dr, x, y, text, f, spacing=0):
    if spacing:
        for ch in text:
            dr.text((x * K, y * K), ch, font=f, fill=0)
            x += dr.textlength(ch, font=f) / K + spacing
        return
    dr.text((x * K, y * K), text, font=f, fill=0)


WORDS = ("charged particle beam source plasma chamber vessel coil grid aperture uniformity divergence substrate ion optics "
         "flux density meniscus sheath potential electrode surface process apparatus magnetic field frequency winding").split()


def paragraph(dr, x, y, width, number, lines, f, pitch, seed):
    """一段两端对齐的正文：首行是段落编号 + 很宽的空 + 正文；返回下一段的起点"""
    k = seed
    for n in range(lines):
        lx = x
        if n == 0 and number:
            ink(dr, lx, y, f"[{number:04d}]", font(SERIF_BOLD, 8.5))
            lx += 34
        words = []
        space = dr.textlength(" ", font=f) / K
        while True:
            w = WORDS[k % len(WORDS)]
            need = sum(dr.textlength(v, font=f) / K for v in words + [w]) + space * len(words)
            if need > (x + width - lx) and words:
                break
            words.append(w)
            k += 3
        last = n == lines - 1
        if last:
            words = words[: max(2, len(words) // 2)]
        text_w = sum(dr.textlength(v, font=f) for v in words) / K
        gap = 3.0 if last or len(words) < 2 else (x + width - lx - text_w) / (len(words) - 1)
        for v in words:
            ink(dr, lx, y, v, f)
            lx += dr.textlength(v, font=f) / K + gap
        y += pitch
    return y


def text_page(w=612, h=792):
    im = Image.new("L", (round(w * K), round(h * K)), 255)
    dr = ImageDraw.Draw(im)
    body, pitch = font(SERIF, 8.5), 10.2
    ink(dr, 72, 50, "US 2099/0000001 A1", font(SERIF, 10))
    ink(dr, 470, 50, "Jan. 1, 2099", font(SERIF, 10))
    ink(dr, 303, 62, "1", font(SERIF, 10))
    left, right, col = 72, 318, 222
    y = 90
    for k, n in enumerate((1, 2, 3)):
        y = paragraph(dr, left, y, col, n, 6, body, pitch, 7 * k) + 5
    # 居中的全大写小标题：字母挨着（扫描件里连成一个连通块）
    ink(dr, 118, y + 4, "SUMMARY OF THE INVENTION", font(SERIF, 9), spacing=-0.6)
    y += 22
    for k, n in enumerate((4, 5, 6)):
        y = paragraph(dr, left, y, col, n, 7, body, pitch, 11 * k + 5) + 5
    y = 90
    for k, n in enumerate((7, 8)):
        y = paragraph(dr, right, y, col, n, 8, body, pitch, 13 * k + 2) + 5
    # 右栏下面的附图：外框里一行标注，框外的附图标记连着引线
    fx0, fy0, fx1, fy1 = 340, 430, 520, 560
    dr.rectangle([fx0 * K, fy0 * K, fx1 * K, fy1 * K], outline=0, width=3)
    dr.rectangle([360 * K, 450 * K, 430 * K, 490 * K], outline=0, width=3)
    dr.rectangle([450 * K, 450 * K, 500 * K, 530 * K], outline=0, width=3)
    ink(dr, 364, 465, "MATCHING", font(SANS, 7))
    for x, label in ((372, "10"), (468, "12")):
        dr.line([x * K, 452 * K, (x - 8) * K, 590 * K], fill=0, width=1)
        ink(dr, x - 14, 592, label, font(SANS, 8))
    ink(dr, 400, 610, "FIG. 1", font(SANS, 11))
    return bw(im)


def drawing(w, h, fig, header=True):
    """一页附图（按显示方向画）：几个框、连线、附图标记；横向页面页边一条转了 90° 的页眉"""
    im = Image.new("L", (round(w * K), round(h * K)), 255)
    dr = ImageDraw.Draw(im)
    dr.rectangle([120 * K, 120 * K, (w - 160) * K, (h - 160) * K], outline=0, width=4)
    for k in range(3):
        x = 160 + k * (w - 360) / 3
        dr.rectangle([x * K, 200 * K, (x + 60) * K, 300 * K], outline=0, width=3)
        dr.line([(x + 30) * K, 300 * K, (x + 30) * K, (h - 200) * K], fill=0, width=3)
        ink(dr, x + 20, 180, f"{100 + 2 * k}a", font(SANS, 9))
    ink(dr, 140, h - 140, fig, font(SANS, 12))
    if header:
        strip = Image.new("L", (round(420 * K), round(14 * K)), 255)
        ImageDraw.Draw(strip).text((0, 0), "Patent Application Publication  Jan. 1, 2099  Sheet 2 of 3", font=font(SERIF, 9), fill=0)
        im.paste(strip.rotate(-90, expand=True), (round((w - 40) * K), round(100 * K)))
    return bw(im)


def table_page(w=612, h=792):
    im = Image.new("L", (round(w * K), round(h * K)), 255)
    dr = ImageDraw.Draw(im)
    xs, ys = (90, 300, 400, 500), (120, 170, 230, 290)
    for x in xs:
        dr.line([x * K, ys[0] * K, x * K, ys[-1] * K], fill=0, width=5)
    for y in ys:
        dr.line([xs[0] * K, y * K, xs[-1] * K, y * K], fill=0, width=5)
    f = font(SANS, 10)
    for r, (name, a, b) in enumerate((("Source Type", "Angle", "Change"), ("350 mm Production Source", "7.7", "0.8"), ("450 mm Source with Grid", "2.4", "0.5"))):
        y = ys[r] + 18
        ink(dr, xs[0] + 10, y, name, f)
        ink(dr, xs[1] + 20, y, a, f)
        ink(dr, xs[2] + 20, y, b, f)
    ink(dr, 90, 320, "FIG. 4", font(SANS, 12))
    return bw(im)


def raw_for_quarter_turn(display):
    """转 90° 放置时位图里存的样子：放置矩阵 cm 0 -H W 0 0 H 把位图的列画成页面的纵向、行画成横向"""
    return display.transpose(Image.Transpose.ROTATE_90)


draft = io.BytesIO()
c = canvas.Canvas(draft, pagesize=(612, 792), pageCompression=1, invariant=1)
c.setCreator("doc-agent test fixture")


def put(img, w, h):
    buf = io.BytesIO()
    img.save(buf, "PNG")
    buf.seek(0)
    c.drawImage(ImageReader(buf), 0, 0, w, h)


# 1：文字页
c.setPageSize((612, 792))
put(text_page(), 612, 792)
c.showPage()
# 2：横向页面，附图扫成竖幅、转 90° 放：单位方块 (u, v) → (792 v, 612 (1 − u))
c.setPageSize((792, 612))
c.saveState()
c.transform(0, -612, 792, 0, 0, 612)
put(raw_for_quarter_turn(drawing(792, 612, "FIG. 2")), 1, 1)
c.restoreState()
c.showPage()
# 3：页面框是纵向的 612 × 792，图却按横向 792 × 612 转 90° 放（页面框与图横竖互换），只在页面的下部、右边被裁掉
c.setPageSize((612, 792))
c.saveState()
c.transform(0, -612, 792, 0, 0, 612)
put(raw_for_quarter_turn(drawing(792, 612, "FIG. 3", header=False)), 1, 1)
c.restoreState()
c.showPage()
# 4：表格页
c.setPageSize((612, 792))
put(table_page(), 612, 792)
c.showPage()
c.save()

# reportlab 只会写 8 位灰度图；真正的扫描件是 1 位黑白图（CCITT），逐张改写
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
with open(out, "wb") as fh:
    writer.write(fh)
