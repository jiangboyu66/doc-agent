"""生成测试夹具 colstart.pdf：双栏论文里"图页"的结构。
第 1 页整页双栏正文；第 2 页是通栏大图 + 通栏图题，页面下方只剩 3 行双栏正文（本页的行太少，单页检测不出分栏）；第 3 页又是整页双栏。"""
import sys
from PIL import Image, ImageDraw
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

W, H = 576, 783
out = sys.argv[1] if len(sys.argv) > 1 else "colstart.pdf"
FS, LEAD = 9.5, 11.0
LX, CW, RX = 36, 242, 297
WORDS = ("segmentation of medical images with transformer encoders reduces the size of the original images while the decoder keeps spatial detail "
         "and the performance of the model slightly decreased on several datasets when the input resolution was reduced to 128 by 128 for training ").split()

c = canvas.Canvas(out, pagesize=(W, H))

def stream(n):
    i = 0
    while True:
        yield WORDS[(i + n) % len(WORDS)]; i += 1

def justified_lines(gen, width, count):
    lines = []
    for _ in range(count):
        cur = []
        while True:
            w = next(gen)
            if c.stringWidth(" ".join(cur + [w]), "Times-Roman", FS) > width and cur: gen_back.append(w); break
            cur.append(w)
        lines.append(cur)
    return lines

gen_back = []
def lines_for(width, count, seed):
    g = stream(seed); out_lines = []
    for _ in range(count):
        cur = []
        while True:
            w = gen_back.pop() if gen_back else next(g)
            if cur and c.stringWidth(" ".join(cur + [w]), "Times-Roman", FS) > width:
                gen_back.append(w); break
            cur.append(w)
        out_lines.append(cur)
    gen_back.clear()
    return out_lines

def draw_col(x, top, width, count, seed):
    for k, words in enumerate(lines_for(width, count, seed)):
        t = c.beginText(x, H - top - k * LEAD)
        t.setFont("Times-Roman", FS)
        natural = c.stringWidth(" ".join(words), "Times-Roman", FS)
        t.setWordSpace((width - natural) / max(1, len(words) - 1))
        t.textLine(" ".join(words)); c.drawText(t)

def header(n):
    c.setFont("Helvetica", 7); c.drawString(36, H - 31, "T.-H. Pham et al.: seUNet-Trans: A Simple Yet Effective UNet-Transformer Model")
    c.setLineWidth(0.4); c.line(36, H - 37, 538, H - 37); c.drawString(36, 38, "VOLUME 12, 2024")

def full_page(seed):
    header(0)
    draw_col(LX, 70, CW, 58, seed); draw_col(RX, 70, CW, 58, seed + 3)
    c.showPage()

full_page(0)

# 第 2 页：通栏大图 + 通栏图题 + 页面下方 3 行双栏正文
header(1)
im = Image.new("RGB", (800, 400))
d = ImageDraw.Draw(im)
for y in range(400): d.line([(0, y), (800, y)], fill=(180 - y // 4, 120 + y // 5, 200))
c.drawImage(ImageReader(im), 36, H - 70 - 255, 502, 251)
c.setFont("Helvetica-Bold", 7.5)
c.drawString(36, H - 345, "FIGURE 11.")
c.setFont("Helvetica", 7.5)
c.drawString(36 + c.stringWidth("FIGURE 11. ", "Helvetica-Bold", 7.5), H - 345, "Qualitative comparison of different predictions on different datasets by visualization. On a specific set of images")
c.drawString(36, H - 355, "covered by a red rectangular, the input, prediction, and ground are organized from the left to the right.")
draw_col(LX, 385, CW, 3, 2); draw_col(RX, 385, CW, 3, 5)
c.showPage()

full_page(7)
c.save()
