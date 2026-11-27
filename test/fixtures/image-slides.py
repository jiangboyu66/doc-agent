"""生成测试夹具 image-slides.pdf：16:9 幻灯片，第一页只有一行居中的标题，后两页各是一张横跨整页的插图（没有文字）。
按正文统计的版心只有标题那么宽；以前可编辑排版把插图的左边只许伸出一点，整张图被挤到右边、超出页面的部分被裁掉，
报告还说"每页保留为一张整页图片、可能是扫描件"。"""
import io
import sys
from PIL import Image, ImageDraw
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

out = sys.argv[1] if len(sys.argv) > 1 else "image-slides.pdf"
W, H = 720, 405
c = canvas.Canvas(out, pagesize=(W, H))

c.setFont("Helvetica-Bold", 28)
c.drawCentredString(W / 2, H / 2, "Spring Picnic")
c.showPage()


def picture(seed):
    im = Image.new("RGB", (620, 330), (120 + seed * 40, 180, 230 - seed * 50))
    d = ImageDraw.Draw(im)
    for k in range(12):
        x = (k * 53 + seed * 17) % 600
        d.ellipse([x, 40 + (k * 37) % 220, x + 60, 100 + (k * 37) % 220], fill=(250, 220 - k * 10, 80 + seed * 30))
    d.rectangle([0, 280, 620, 330], fill=(90, 160, 70))
    buf = io.BytesIO()
    im.save(buf, "PNG")
    buf.seek(0)
    return ImageReader(buf)


for seed in (1, 2):
    c.drawImage(picture(seed), 50, 40, width=620, height=330)
    c.showPage()
c.save()
