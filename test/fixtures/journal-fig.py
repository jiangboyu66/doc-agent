"""生成测试夹具 journal-fig.pdf（IEEE 类期刊的两页，双栏）：
  - 第 1 页：作者名后夹着矢量小图标（ORCID：几条绿色填充的曲线路径，不是位图）；
  - 第 2 页：通栏上下叠放的两幅图——每幅由左右两张相邻的位图拼成（3 列 + 2 列），图内有白色小标签，
            两幅图之间只隔着第一幅的图题；图题之后是双栏正文。两幅图不能被并成一组（否则图题被夹进去、整组被否决，
            两幅图各自按零散位图、按所在栏拆开，右半幅掉到页底、与页脚重叠，正文被挤到下一页）；
  - 第 3 页：双栏里的多行公式，续行右端带编号、没有关系符号（"−λ log p(…)  (10)"），上一行是带大括号的主式。"""
import sys
from PIL import Image, ImageDraw
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import tempfile, os
from fontTools.ttLib import TTFont as FT

W, H = 612, 792
out = sys.argv[1] if len(sys.argv) > 1 else "journal-fig.pdf"
LIB = "/usr/share/fonts/truetype/liberation/LiberationSerif-%s.ttf"
def renamed(src, name):
    f = FT(src)
    for rec in f["name"].names:
        if rec.nameID in (1, 3, 4, 6): rec.string = name
    path = os.path.join(tempfile.mkdtemp(), name + ".ttf"); f.save(path); return path
for reg, style in (("NimbusRomNo9L-Regu", "Regular"), ("NimbusRomNo9L-ReguItal", "Italic"), ("MTSYN", "Regular"), ("RMTMI", "Italic")):
    pdfmetrics.registerFont(TTFont(reg, renamed(LIB % style, reg)))
c = canvas.Canvas(out, pagesize=(W, H))
y_ = lambda top: H - top
WORDS = ("the proposed method restores the texture of the reconstructed images while the denoising network removes the noise and "
         "keeps the structure of the low dose computed tomography slices in the frequency domain ").split()
FONT = "NimbusRomNo9L-Regu"

def column(x0, x1, top, n, seed, size=10, lead=12):
    k = seed
    for i in range(n):
        cur = []
        while True:
            w = WORDS[k % len(WORDS)]
            if cur and c.stringWidth(" ".join(cur + [w]), FONT, size) > x1 - x0: break
            cur.append(w); k += 1
        t = c.beginText(x0, y_(top + i * lead)); t.setFont(FONT, size)
        t.setWordSpace((x1 - x0 - c.stringWidth(" ".join(cur), FONT, size)) / max(1, len(cur) - 1))
        t.textLine(" ".join(cur)); c.drawText(t)

def text(x, top, s, font=FONT, size=10): c.setFont(font, size); c.drawString(x, y_(top), s); return x + c.stringWidth(s, font, size)

def photo(seed, w, h):
    im = Image.new("RGB", (w, h), (20, 20, 20)); d = ImageDraw.Draw(im)
    d.ellipse((w * .2, h * .15, w * .8, h * .85), fill=(110 + seed * 12, 110, 120)); return ImageReader(im)

def figure(top, labels_left, labels_right, caption):
    """两行 × (3 + 2) 列：左边一张位图装 3 列、右边一张装 2 列；每个面板左上角有白色小标签"""
    h = 90
    for r in range(2):
        y = top + r * (h + 4)
        c.drawImage(photo(r, 400, 160), 62, y_(y + h), 279, h)
        c.drawImage(photo(r + 3, 300, 160), 341, y_(y + h), 209, h)
        for k, lab in enumerate(labels_left[r]): c.setFillColorRGB(1, 1, 1); c.setFont(FONT, 7); c.drawString(64 + k * 93, y_(y + 8), lab)
        for k, lab in enumerate(labels_right[r]): c.setFillColorRGB(1, 1, 1); c.setFont(FONT, 7); c.drawString(343 + k * 104, y_(y + 8), lab)
        c.setFillColorRGB(0, 0, 0)
    text(49, top + 2 * (h + 4) + 10, caption[0], size=8); text(49, top + 2 * (h + 4) + 19, caption[1], size=8)
    return top + 2 * (h + 4) + 19

# ---- 第 1 页：作者行 + 矢量图标
text(150, 80, "Zero-Shot Low-Dose CT Image Denoising", size=20)
x = 90
for name in ("Bo Su", "Xiangyun Hu", "Yunfei Zha", "Baochang Zhang"):
    x = text(x, 120, name, size=10) + 1.5
    c.setFillColorRGB(0.65, 0.81, 0.22)
    p = c.beginPath(); cx, cy, r = x + 4.5, y_(118), 4.5
    p.moveTo(cx + r, cy); p.curveTo(cx + r, cy + r * .55, cx + r * .55, cy + r, cx, cy + r); p.curveTo(cx - r * .55, cy + r, cx - r, cy + r * .55, cx - r, cy)
    p.curveTo(cx - r, cy - r * .55, cx - r * .55, cy - r, cx, cy - r); p.curveTo(cx + r * .55, cy - r, cx + r, cy - r * .55, cx + r, cy); p.close()
    c.drawPath(p, stroke=0, fill=1)
    c.setFillColorRGB(1, 1, 1); c.rect(cx - 1, cy - 2.5, 1.6, 5, stroke=0, fill=1); c.setFillColorRGB(0, 0, 0)
    x += 12 + 4
column(49, 300, 170, 28, 0); column(312, 563, 170, 28, 7)
c.showPage()

# ---- 第 2 页：两幅图上下叠放
nxt = figure(54, [["LDCT", "NLM", "EDCNN"], ["NDCT", "WGAN", "N2S"]], [["REDCNN", "CTformer"], ["SR3", "Ours"]],
             ("Fig. 8.  Axial results of different denoising methods on the Mayo 2016 dataset. The window is [-100, 200] HU.", "Zoomed parts over the region of interest (ROI) marked by the yellow box."))
nxt = figure(nxt + 4, [["LDCT", "NLM", "EDCNN"], ["NDCT", "WGAN", "N2S"]], [["REDCNN", "CTformer"], ["SR3", "Ours"]],
             ("Fig. 9.  Comparison of denoising performance of Mayo data slices in different methods.", "The yellow ROI areas are zoomed in to show the denoising details."))
column(49, 300, nxt + 20, 12, 3); column(312, 563, nxt + 20, 12, 9)
c.showPage()

# ---- 第 3 页：多行公式（续行右端带编号、没有关系符号）
column(49, 300, 60, 8, 2); column(312, 563, 60, 8, 5)
def math(x, top, parts, size=10):
    for font, t in parts: x = text(x, top, t, font, size)
    return x
top = 175
text(130, top + 9, "\u011d", "NimbusRomNo9L-ReguItal", 10)
math(140, top + 10, [("MTSYN", "="), (FONT, " arg min "), ("MTSYN", "\u2225"), ("NimbusRomNo9L-ReguItal", "g"), ("MTSYN", "\u2212"), ("NimbusRomNo9L-ReguItal", "l"), ("MTSYN", "\u2225")])
c.setLineWidth(0.5); c.setFont("MTSYN", 20); c.drawString(122, y_(top + 17), "["); c.drawString(262, y_(top + 28), "]")
math(160, top + 30, [("MTSYN", "\u2212"), ("RMTMI", "\u03bb"), (FONT, " log "), ("NimbusRomNo9L-ReguItal", "p"), ("RMTMI", "(g | g)")])
text(268, top + 30, "(10)", size=10)
column(49, 300, top + 60, 10, 6); column(312, 563, top + 60, 10, 1)
c.showPage()
c.save()
