"""生成测试夹具 texpaper.pdf：LaTeX（pdfTeX）论文首页常见的几种排版。

1. 作者栏：三位作者并排，邮箱用 TeX 的打字机字体 CMTT10（每字 0.525 em），摘要里居中的网址也是 CMTT10。
   换成 Word 的 Courier New（每字 0.6 em）后变宽，并排的邮箱互相挤压、折到下一行；
   以前还把邮箱之间的制表间距算成"加宽字距"，反而给每个字加上字符间距。
2. \\includegraphics[trim, clip]：图片放在带 BBox 的表单 XObject 里，BBox 之外的部分被裁掉。
   以前不处理表单的 BBox，图片按未裁剪的原尺寸放进 Word，比原文大一倍、盖住旁边的图和文字。
3. 示意图：一行圆圈（节点）、节点旁的小图、节点之间是文字符号 "· · ·"。小图用矩形裁剪（路径末尾多一个 moveto，
   pdfTeX 的写法）只露出一部分。以前只有其中一段被识别成图，其余节点和小图散落成文字和图片，排版错乱。
4. 曲线图：纵轴标题是旋转文字（不转换），刻度标签 9 pt（正文 10 pt）。以前图区域没包含它们，
   纵轴标题丢失、刻度数字散落在图下方。

用法：python3 texpaper.py [输出路径]（需要 reportlab、Pillow；CMTT10 取自 TeX Live 的 amsfonts）
"""
import random
import sys

from PIL import Image
from reportlab.lib.colors import Color
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas

out = sys.argv[1] if len(sys.argv) > 1 else "texpaper.pdf"
CM = "/usr/share/texlive/texmf-dist/fonts/"
face = pdfmetrics.EmbeddedType1Face(CM + "afm/public/amsfonts/cm/cmtt10.afm", CM + "type1/public/amsfonts/cm/cmtt10.pfb")
pdfmetrics.registerTypeFace(face)
pdfmetrics.registerFont(pdfmetrics.Font("CMTT10", face.name, "WinAnsiEncoding"))

W, H = 612, 792
c = canvas.Canvas(out, pagesize=(W, H))
random.seed(7)


def img(w, h, seed, kind="noise"):
    """确定性的测试图片：噪声 / 渐变块"""
    rnd = random.Random(seed)
    im = Image.new("RGB", (w, h))
    px = im.load()
    for y in range(h):
        for x in range(w):
            if kind == "noise":
                v = rnd.randint(60, 200)
                px[x, y] = (v, v, v)
            else:
                px[x, y] = ((x * 7 + seed * 40) % 256, (y * 5 + seed * 20) % 256, ((x + y) * 3) % 256)
    return ImageReader(im)


def text(x, y, s, font="Times-Roman", size=10, center=False):
    c.setFont(font, size)
    (c.drawCentredString if center else c.drawString)(x, H - y, s)


# ---- 标题与作者栏 ----
c.setFillColor(Color(0, 0, 0))
c.rect(108, H - 83, 396, 4, stroke=0, fill=1)
text(306, 113, "Denoising Toy Models for Layout Tests", "Times-Bold", 17, center=True)
c.setLineWidth(0.5)
c.line(108, H - 132, 504, H - 132)
for cx, name, mail in [(180, "Alice Zhang", "alicezhang@example.edu"), (306, "Bo Li", "boli@example.edu"), (431, "Carol Smith", "carols@cs.example.edu")]:
    text(cx, 169, name, "Times-Bold", 10, center=True)
    text(cx, 180, "Example University", center=True)
    text(cx, 191, mail, "CMTT10", 10, center=True)

text(306, 232, "Abstract", "Times-Bold", 12, center=True)
abstract = [
    "We present synthetic layout samples that exercise the conversion of figures and author blocks,",
    "a class of test documents inspired by considerations from typical conference papers. Our",
]
for i, s in enumerate(abstract):
    text(306, 256 + 11 * i, s, center=True)
# 末行居中：正文 + 打字机字体的网址
lead, url = "implementation is available at ", "https://example.org/toy/layout"
tw = c.stringWidth(lead, "Times-Roman", 10) + c.stringWidth(url, "CMTT10", 10)
text(306 - tw / 2, 278, lead)
text(306 - tw / 2 + c.stringWidth(lead, "Times-Roman", 10), 278, url, "CMTT10", 10)

# ---- 图 1：左边普通图片，右边 \includegraphics[clip] 裁剪后的图片 ----
c.drawImage(img(64, 64, 1, "grad"), 108, H - 520, 200, 200)
c.beginForm("clipped", 0, 0, 150, 200)  # BBox：只显示 150 × 200
c.drawImage(img(96, 96, 2, "grad"), 0, -100, 300, 300)  # 原图 300 × 300，大部分被 BBox 裁掉
c.endForm()
c.saveState()
c.translate(312, H - 520)
c.doForm("clipped")
c.restoreState()
text(306, 534, "Figure 1: Generated samples (left) and a cropped sample grid (right).", size=9, center=True)

# ---- 正文 ----
body = [
    "This paragraph is ordinary body text in ten point Times so that the converter can determine the body",
    "size of the page. The figures below are drawn the way pdfTeX draws them for a conference paper.",
]
for i, s in enumerate(body):
    text(108, 560 + 12 * i, s)

# ---- 图 2：节点示意图 ----
row = 610


def node(cx, label, fill=True):
    c.setFillColor(Color(0.75, 0.75, 0.75) if fill else Color(1, 1, 1))
    c.circle(cx, H - row, 11.8, stroke=1, fill=1)
    c.setFillColor(Color(0, 0, 0))
    text(cx - 4, row + 3.5, label, "Times-Bold", 12)


def thumb(x, seed):
    """节点右下角的小图：原图 40 × 40，按 pdfTeX 的写法（路径末尾多一个 moveto）裁剪成右下角 20 × 20"""
    yb = H - row - 20
    c._code.append(f"q {x:.2f} {yb:.2f} m {x + 20:.2f} {yb:.2f} l {x + 20:.2f} {yb + 20:.2f} l {x:.2f} {yb + 20:.2f} l h {x:.2f} {yb:.2f} m W n")
    c.drawImage(img(32, 32, seed), x - 20, yb, 40, 40)
    c._code.append("Q")


node(170, "x")
thumb(182, 3)
text(212, row + 3.5, "· · ·", size=12)
node(259, "y")
c.setDash(2, 2)
c.bezier(282, H - row - 9, 292, H - row - 18, 306, H - row - 18, 318, H - row - 9)
c.setDash()
node(331, "z")
thumb(343, 4)
text(373, row + 3.5, "· · ·", size=12)
node(420, "w", fill=False)
thumb(432, 5)
text(306, 640, "Figure 2: The directed graphical model considered in this test.", size=9, center=True)

# ---- 图 3：曲线图（旋转的纵轴标题、9 pt 刻度） ----
px0, py0, pw, ph = 170, 668, 140, 70  # 版面坐标（y 向下）
c.setStrokeColor(Color(0.75, 0.75, 0.75))
c.setLineWidth(0.4)
for k in range(1, 7):
    c.line(px0 + k * pw / 7, H - py0, px0 + k * pw / 7, H - py0 - ph)
for k in range(1, 4):
    c.line(px0, H - py0 - k * ph / 4, px0 + pw, H - py0 - k * ph / 4)
c.setStrokeColor(Color(0, 0, 0))
c.rect(px0, H - py0 - ph, pw, ph, stroke=1, fill=0)
c.setFillColor(Color(0, 0, 0.8))
for k in range(12):
    c.circle(px0 + 8 + k * 11, H - py0 - ph + 8 + (k * k) * 0.45, 1.8, stroke=0, fill=1)
c.setFillColor(Color(0, 0, 0))
for k, v in enumerate(["2", "4", "6", "8"]):
    text(px0 - 7, py0 + ph - k * ph / 4 + 3 - 2, v, size=9)
for k, v in enumerate(["0", "200", "400", "600"]):
    text(px0 + k * pw / 3.2, py0 + ph + 10, v, size=9)
c.saveState()
c.translate(px0 - 14, H - py0 - ph + 12)
c.rotate(90)
c.setFont("Times-Roman", 9)
c.drawString(0, 0, "Inception Score")
c.restoreState()
text(306, 768, "Figure 3: Sample quality over time.", size=9, center=True)

c.showPage()
c.save()
