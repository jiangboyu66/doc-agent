"""生成测试夹具 transparent-figure.pdf：两栏正文，左栏顶上一张带透明通道的图（Office 图表导出的样子）：
不透明的图表只占左边 200pt，右边 100pt 是完全透明的边，整张图按原尺寸会伸进右栏。
以前按整张图的范围排版：图被当成横跨两栏的图，两栏正文整体往下挤到图下面，栏底的行溢到第 2 页。"""
import io
import struct
import sys
import zlib
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

out = sys.argv[1] if len(sys.argv) > 1 else "transparent-figure.pdf"
W, H = 612, 792
c = canvas.Canvas(out, pagesize=(W, H))


def y(v):
    return H - v


def rgba_png(w, h, opaque_w):
    """左边 opaque_w 列是不透明的图表（浅蓝底、深蓝折线），右边完全透明"""
    rows = []
    for r in range(h):
        row = bytearray(b"\x00")
        for x in range(w):
            if x < opaque_w:
                on_line = abs((h - 1 - r) - (x * (h - 20) // opaque_w + 10) // 2 - h // 4) < 2
                row += bytes((30, 60, 160, 255)) if on_line else bytes((225, 235, 250, 255))
            else:
                row += bytes((0, 0, 0, 0))
        rows.append(bytes(row))
    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    data = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(b"".join(rows))) + chunk(b"IEND", b"")
    return ImageReader(io.BytesIO(data))


L0, L1, R0, R1 = 54, 297, 315, 558
c.setFont("Times-Roman", 10)
words = ("The converter keeps each column separate and preserves the line breaks of the original "
         "document so that the pages line up with the source when they are compared side by side ").split()


def column(x0, x1, top, bottom):
    line, base, k = "", top, 0
    while base <= bottom:
        w = words[k % len(words)]
        k += 1
        t = (line + " " + w).strip()
        if c.stringWidth(t, "Times-Roman", 10) > x1 - x0:
            c.drawString(x0, y(base), line)
            base += 12
            line = w
        else:
            line = t


# 左栏顶上的图：300 × 150pt，左边 200pt 不透明，右边 100pt 透明（伸进右栏）
c.drawImage(rgba_png(300, 150, 200), L0, y(72 + 150), 300, 150, mask="auto")
c.setFont("Times-Italic", 9)
c.drawString(L0, y(236), "Fig. 1. A chart exported with a transparent margin.")
c.setFont("Times-Roman", 10)
column(L0, L1, 256, 740)
column(R0, R1, 72, 740)
c.showPage()
c.save()
