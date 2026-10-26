"""生成测试夹具 journal-eq.pdf：Elsevier 类期刊的另一些结构（三页，双栏）。
  - 第 1 页：刊头灰底用一根 60 磅宽的描边线画成（不是色块），刊名叠在上面；自定义字体名（Gulliver）的正文；
            逐字母定位的加宽字距标题（每个字母一个文字项）。
  - 第 2 页：偶数页，右栏从 302 起；栏内分式公式：整行变量（"(2μ + c)(2σ + c)"）在分数线上方、分母在下方，
            没有编号的公式在栏内、编号在栏右端；变量上方的重音（ȳ）是独立文字项。
  - 第 3 页：奇数页，右栏从 312 起、整页版心比第 2 页偏右 10 磅；公式 "LDCT = iradon ln (I0 / (I0 × exp(…)))" 分数线落在行内。"""
import sys
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

W, H = 595.3, 793.7
out = sys.argv[1] if len(sys.argv) > 1 else "journal-eq.pdf"
import io, tempfile, os
from fontTools.ttLib import TTFont as FT

def renamed(src, ps_name):
    """把 Liberation Serif 复制一份并改写字体名（reportlab 嵌入时用字体自己的 PostScript 名，而不是注册名）"""
    f = FT(src)
    for rec in f["name"].names:
        if rec.nameID in (1, 3, 4, 6): rec.string = ps_name if rec.nameID != 4 else ps_name
    path = os.path.join(tempfile.mkdtemp(), ps_name + ".ttf"); f.save(path); return path

LIB = "/usr/share/fonts/truetype/liberation/LiberationSerif-%s.ttf"
# Gulliver：Elsevier 的正文字体（Word 里没有）；MTSY / RMTMI：出版社的数学字体（MathTime），转换器按字体名识别为数学字体
for reg, style, ps in (("Gulliver", "Regular", "Gulliver"), ("Gulliver-Italic", "Italic", "Gulliver-Italic"), ("MTSY", "Regular", "MTSY"), ("RMTMI", "Italic", "RMTMI")):
    pdfmetrics.registerFont(TTFont(reg, renamed(LIB % style, ps)))
c = canvas.Canvas(out, pagesize=(W, H))
y_ = lambda top: H - top
WORDS = ("the network learns to restore the texture of the reconstructed images in the frequency domain with a dedicated loss and "
         "low dose computed tomography reduces the radiation exposure of patients while the image noise increases ").split()

def column(x0, x1, top, n, seed, size=8, lead=10.45, font="Gulliver"):
    k = seed
    for i in range(n):
        cur = []
        while True:
            w = WORDS[k % len(WORDS)]
            if cur and c.stringWidth(" ".join(cur + [w]), font, size) > x1 - x0: break
            cur.append(w); k += 1
        t = c.beginText(x0, y_(top + i * lead)); t.setFont(font, size)
        t.setWordSpace((x1 - x0 - c.stringWidth(" ".join(cur), font, size)) / max(1, len(cur) - 1))
        t.textLine(" ".join(cur)); c.drawText(t)

def text(x, top, s, font="Gulliver", size=8):
    c.setFont(font, size); c.drawString(x, y_(top), s); return x + c.stringWidth(s, font, size)

# ---- 第 1 页
c.setStrokeColorRGB(0.9, 0.9, 0.9); c.setLineWidth(65); c.line(116.6, y_(96.3), 493.6, y_(96.3))
c.setStrokeColorRGB(0, 0, 0); c.setLineWidth(0.25); c.line(42.8, y_(57.3), 493.5, y_(57.3))
text(200, 100, "Computer Methods and Programs in Biomedicine", size=12)
x = 42.8
for ch in "ARTICLE":
    x = text(x, 160, ch, size=9) + 3.4
column(42.8, 284, 190, 40, 0); column(312, 553, 190, 40, 9)
c.showPage()

# ---- 第 2 页（偶数页：右栏 302 起）
column(33, 284, 60, 20, 4); column(302, 553, 60, 20, 8)
c.setLineWidth(0.53)
def math(x, top, parts, size=8):
    """parts: [(字体, 文字)]；数学字体的片段按出版社的做法单独成文字项"""
    for font, t in parts: x = text(x, top, t, font, size)
    return x
text(302, 300, "SSIM", "Gulliver-Italic", 8.5); math(330, 300, [("MTSY", "=")], 8.5)
math(425, 291.5, [("RMTMI", "("), ("Gulliver", "2"), ("RMTMI", "\u03bc"), ("MTSY", " + "), ("Gulliver-Italic", "c"), ("RMTMI", ")("), ("Gulliver", "2"), ("RMTMI", "\u03c3"), ("MTSY", " + "), ("Gulliver-Italic", "c"), ("RMTMI", ")")])
c.line(342, y_(296.6), 468, y_(296.6))
math(345, 308, [("RMTMI", "("), ("RMTMI", "\u03bc"), ("MTSY", " + "), ("RMTMI", "\u03bc"), ("MTSY", " + "), ("Gulliver-Italic", "c"), ("RMTMI", ")("), ("RMTMI", "\u03c3"), ("MTSY", " + "), ("RMTMI", "\u03c3"), ("MTSY", " + "), ("Gulliver-Italic", "c"), ("RMTMI", ")")])
text(540, 300, "(10)", size=8.5)
text(302, 330, "where", size=8); y0 = 340
text(302, y0, "y", "Gulliver-Italic", 8); text(309, y0 - 4.5, "¯", size=8)
column(302, 553, 360, 8, 3)
column(33, 284, 300, 22, 5)
c.showPage()

# ---- 第 3 页（奇数页：右栏 312 起）
column(33, 284, 60, 20, 6); column(312, 563, 60, 20, 2)
text(312, 280, "LDCT", size=8); text(333, 281.5, "simulated", size=5.5); text(358, 280, "iradon ln", "Gulliver-Italic", 8); math(340, 280, [("MTSY", "=")])
text(458, 271, "I", "Gulliver-Italic", 8)
c.line(418, y_(276.6), 506, y_(276.6))
math(420, 287, [("Gulliver-Italic", "I"), ("MTSY", " \u00d7 "), ("Gulliver", "exp"), ("RMTMI", "("), ("MTSY", "\u2212"), ("Gulliver", "radon"), ("RMTMI", "("), ("Gulliver-Italic", "y"), ("RMTMI", "))"), ("MTSY", " + "), ("Gulliver-Italic", "P"), ("RMTMI", "("), ("RMTMI", "\u03bb"), ("RMTMI", ")")])
text(545, 280, "(8)", size=8.5)
column(312, 563, 310, 6, 1)
# 同一基线上的四个粗体 R，每个后面跟着叠放的上下标和逗号（R₁¹, R₂¹, …）：不是"加宽字距的词"，不能并成 "RRRR"
x = text(33, 280, "the 4-channel images ", size=8)
for k in range(1, 5):
    x = text(x, 280, "R", "Gulliver", 8)
    text(x, 277, "1", size=6); text(x, 283, str(k), size=6); x += 4
    x = text(x, 280, ", " if k < 4 else "", "RMTMI", 8) + 3
c.showPage()
c.save()
