"""生成测试夹具 logo-corner.pdf：绘本式幻灯片（横向）。
第 1 页：横贯整页的深红色横幅，上面一行 40pt 的白色标题（全文只有这一行字），页面右下角一个小徽标（伸进右边距几磅）；
第 2 页：一整页的插图。
以前可编辑排版把徽标挤到了第 2 页、插图排到第 3 页，后面各页整体错后一页：
  - 全文的"正文字号"就是标题的 40pt，写成了文档的默认字号；LibreOffice 按默认字号算只有一张图的段落的行高，
    32pt 高的徽标那一段被撑到 60 多磅高，放不下；
  - 徽标伸出版心几磅，图片段落没有负的右缩进，这一行也放不下它。"""
import io
import sys
import struct
import zlib
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

out = sys.argv[1] if len(sys.argv) > 1 else "logo-corner.pdf"
W, H = 720, 540
c = canvas.Canvas(out, pagesize=(W, H))


def y(v):
    return H - v


def png(w, h, color):
    """纯色 PNG（不依赖 PIL）"""
    r, g, b = color
    raw = b"".join(b"\x00" + bytes([r, g, b]) * w for _ in range(h))
    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    data = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")
    return ImageReader(io.BytesIO(data))


# 第 1 页
c.setFillColorRGB(0.65, 0.17, 0.18)
c.rect(0, y(330), W, 180, stroke=0, fill=1)
c.setFillColorRGB(1, 1, 1)
c.setFont("Helvetica-Bold", 40)
c.drawCentredString(W / 2, y(252), "Peter and the Wolf")
# 右下角的徽标：x 615–716，y 489–521（伸进右边距）
c.drawImage(png(64, 20, (40, 110, 170)), 615, y(521), 101, 32)
c.showPage()

# 第 2 页：一整页的插图（留一点白边）
c.drawImage(png(96, 72, (120, 170, 90)), 12, y(528), 696, 516)
c.showPage()
c.save()
