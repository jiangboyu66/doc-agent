"""生成测试夹具 scan-letter.pdf：3 页扫描信件（每页一张整页的扫描图，A4）。
  - 第 1 页：只有扫描图（多数扫描仪的输出）；
  - 第 2 页：扫描图 + 不可见的 OCR 识别文字层（文字渲染模式 3，"可搜索 PDF"）；
  - 第 3 页：扫描仪把一页切成上下两条图片，拼在一起铺满整页。
易错点：整页图片当成"独占一段的行内图片"，比版心高——逐页保留模式下页首的分页锚点之后放不下，被挤到下一页，
3 页变 6 页、空白页与图片页交替；可编辑排版下图片从上边距处开始排、下半截被裁掉；
OCR 识别层被转成看得见的正文，与图里的字重叠成两层。"""
import io, random, sys
from PIL import Image, ImageDraw
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

W, H = 595.3, 841.9
out = sys.argv[1] if len(sys.argv) > 1 else "scan-letter.pdf"
R = random.Random(3)
PX = 1  # 每 pt 的像素数（72 dpi，夹具保持小巧）
LINES = ["致：上海某某贸易有限公司", "关于《分配协议》履行情况的函", "贵司于2026年10月15日发出的通知已收悉，现函复如下："]


def scan_page(seed):
    """仿扫描页：浅灰底，一行行"文字"（深色短块）"""
    r = random.Random(seed)
    w, h = int(W * PX), int(H * PX)
    im = Image.new("L", (w, h), 244)
    d = ImageDraw.Draw(im)
    y = 72 * PX
    while y < h - 100 * PX:
        x = 72 * PX + (24 * PX if r.random() < 0.2 else 0)
        end = w - 72 * PX if r.random() < 0.8 else r.randint(w // 3, w - 72 * PX)
        while x < end:
            cw = r.randint(9, 12) * PX
            d.rectangle([x, y, x + cw - 2 * PX, y + 10 * PX], fill=50)
            x += cw + PX
        y += 22 * PX
    return im


def jpeg(im):
    """第 1 页按扫描仪的习惯存成 JPEG；其余页用无损压缩（夹具小巧，对版面识别没有区别）"""
    b = io.BytesIO(); im.save(b, "JPEG", quality=35); b.seek(0); return ImageReader(b)


def flate(im):
    return ImageReader(im)


pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
c = canvas.Canvas(out, pagesize=(W, H))
# 第 1 页：只有扫描图
c.drawImage(jpeg(scan_page(1)), 0, 0, W, H); c.showPage()
# 第 2 页：扫描图 + 不可见的 OCR 文字层
c.drawImage(flate(scan_page(2)), 0, 0, W, H)
for k, s in enumerate(LINES * 6):
    t = c.beginText(72, H - 82 - 22 * k); t.setTextRenderMode(3); t.setFont("STSong-Light", 10.5); t.textOut(s); c.drawText(t)
c.showPage()
# 第 3 页：上下两条拼成整页
im = scan_page(3)
top, bottom = im.crop((0, 0, im.width, im.height // 2)), im.crop((0, im.height // 2, im.width, im.height))
c.drawImage(flate(top), 0, H / 2, W, H / 2); c.drawImage(flate(bottom), 0, 0, W, H / 2); c.showPage()
c.save()
