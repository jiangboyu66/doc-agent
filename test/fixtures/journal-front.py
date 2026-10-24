"""生成测试夹具 journal-front.pdf：Elsevier 类期刊的版式结构（两页，双栏）。
  - 第 1 页：作者名后夹着小图标（ORCID）；"ARTICLE INFO"（关键词）与"ABSTRACT"左右并排，各自标题下有一条横线，
    标题用加宽的字距排；下面是双栏正文。
  - 第 2 页：左栏正文；右栏上下两张三线表（横线左右端点几乎相同），中间夹着第二张表的表题；
    右栏底部是带分数线的公式（没有编号、不缩进），紧挨着下面的页码。"""
import sys
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

W, H = 595.3, 793.7
L0, L1, R0, R1 = 37.9, 289.0, 306.9, 557.9
out = sys.argv[1] if len(sys.argv) > 1 else "journal-front.pdf"
c = canvas.Canvas(out, pagesize=(W, H))
y_ = lambda top: H - top
WORDS = ("low dose computed tomography reduces the radiation exposure of patients while the image noise increases and the "
         "network learns to restore the texture of the reconstructed images in the frequency domain with a dedicated loss ").split()

def lines_for(width, n, seed, font="Times-Roman", size=8):
    out_, k = [], seed
    for _ in range(n):
        cur = []
        while True:
            w = WORDS[k % len(WORDS)]
            if cur and c.stringWidth(" ".join(cur + [w]), font, size) > width: break
            cur.append(w); k += 1
        out_.append(cur)
    return out_

def column(x0, x1, top, n, seed, size=8, lead=10.45):
    for i, words in enumerate(lines_for(x1 - x0, n, seed, size=size)):
        t = c.beginText(x0, y_(top + i * lead)); t.setFont("Times-Roman", size)
        if i < n - 1: t.setWordSpace((x1 - x0 - c.stringWidth(" ".join(words), "Times-Roman", size)) / max(1, len(words) - 1))
        t.textLine(" ".join(words)); c.drawText(t)

icon = Image.new("RGB", (40, 40), (166, 206, 57))
def text(x, top, s, font="Times-Roman", size=8):
    c.setFont(font, size); c.drawString(x, y_(top), s); return x + c.stringWidth(s, font, size)

# ---- 第 1 页
text(L0, 70, "Low-dose CT imaging via noise intensity field aware networks", size=13.5)
x = L0
for name, sup in (("Zihui Zhao", "a,b"), ("Yanxin Wang", "a"), ("Suqing Tian", "c"), ("Wei Zhao", "a,*")):
    x = text(x, 95, name, size=10.5) + 1.2
    c.drawImage(ImageReader(icon), x, y_(95), 10.8, 10.8); x += 12
    x = text(x, 91.5, sup, size=7.5) + 2
    if name != "Wei Zhao": x = text(x - 2, 95, ", ", size=10.5)
for i, aff in enumerate(["School of Physics, Beihang University, Beijing, China", "Department of Radiation Oncology, Peking University, Beijing, China", "Hangzhou International Innovation Institute, Hangzhou, China"]):
    text(L0, 110 + 8 * i, aff, "Times-Italic", 6.5)
c.setLineWidth(0.4); c.line(L0, y_(150), R1, y_(150))
# 字距加宽的两种写法：字符间距参数（Tc），或逐个字母定位（文字提取时字母之间被插入空格）
t = c.beginText(L0, y_(172)); t.setFont("Times-Roman", 9); t.setCharSpace(3.4); t.textLine("ARTICLE INFO"); t.setCharSpace(0); c.drawText(t)
x = 202.3
for ch in "ABSTRACT":
    x = text(x, 172, ch, size=9) + 3.4
c.line(L0, y_(181.7), 171.1, y_(181.7)); c.line(202.3, y_(181.7), 556.6, y_(181.7))
text(L0, 192, "Keywords:", "Times-Italic", 6.5)
for i, k in enumerate(["Deep learning", "Noise intensity field", "Fourier domain", "Noise power spectrum", "Imaging physics"]):
    text(L0, 200.5 + 8.5 * i, k, size=6.5)
column(202.3, 556.6, 193, 10, 3, size=7, lead=10.1)
c.line(L0, y_(305), R1, y_(305))
column(L0, L1, 330, 40, 0); column(R0, R1, 330, 40, 9)
text(L0, 760, "https://doi.org/10.1016/j.media.2025.103866", size=6.5)
c.showPage()

# ---- 第 2 页
column(L0, L1, 60, 66, 4)
def table(top, cap, seed):
    text(R0 + 3.6, top, cap, "Times-Bold", 7)
    text(R0 + 3.6, top + 9.6, "Quantitative evaluation of the proposed method and other baseline methods.", size=7)
    t0 = top + 15
    c.setLineWidth(0.4); c.line(R0 + 4.2, y_(t0), R1 - 4.2, y_(t0))
    cols = [R0 + 7, R0 + 105, R0 + 130, R0 + 155, R0 + 180, R0 + 210]
    for x0, h in zip(cols, ["Method", "SSIM", "MSE", "PSNR", "LFD", "VIF"]): text(x0, t0 + 9, h, size=6.5)
    c.setLineWidth(0.5); c.line(R0 + 4.2, y_(t0 + 13.3), R1 - 4.2, y_(t0 + 13.3))
    for r, m in enumerate(["Low Dose", "U-Net", "DnCNN", "REDCNN", "SWINIR", "NIFA"]):
        text(cols[0], t0 + 24 + 8.6 * r, m, "Times-Bold" if m == "NIFA" else "Times-Roman", 6.5)
        for k, x0 in enumerate(cols[1:]):
            text(x0, t0 + 24 + 8.6 * r, "%.2f" % (0.5 + ((seed * 7 + r * 3 + k * 5) % 40) / 10), "Times-Bold" if m == "NIFA" else "Times-Roman", 6.5)
    c.line(R0 + 4.2, y_(t0 + 69.7), R1 - 4.2, y_(t0 + 69.7))
    return t0 + 70
b1 = table(60, "Table 1", 1)
b2 = table(b1 + 30, "Table 2", 2)
column(R0, R1, b2 + 25, 30, 6)
eq = b2 + 25 + 30 * 10.45 + 10
text(R0, eq + 4.6, "x", "Times-Italic", 8); text(R0 + 8, eq + 4.6, "= i −", size=8)
text(R0 + 30, eq - 0.3, "W", "Times-Italic", 8); c.setLineWidth(0.53); c.line(R0 + 29.5, y_(eq + 2.1), R0 + 38.5, y_(eq + 2.1)); text(R0 + 32, eq + 9.9, "2", size=8)
text(R0 + 42, eq + 4.6, ", y = j −", "Times-Italic", 8)
text(R0 + 80, eq - 0.3, "H", "Times-Italic", 8); c.line(R0 + 79.5, y_(eq + 2.1), R0 + 87.5, y_(eq + 2.1)); text(R0 + 81, eq + 9.9, "2", size=8)
text(W / 2, eq + 21, "2", size=6.5)
c.showPage()
c.save()
