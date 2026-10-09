"""生成测试夹具：正文用 Times-Roman；公式与符号用"没有文字编码"的字形：
   - 页面内嵌的 1 位图像蒙版（老式 TeX/DVI 生成的 PDF 中常见的位图字形）；
   - Type 3 字体（矢量字形，字符编码为控制字符 \\x01，文字提取得到乱码）。"""
import sys, zlib

def glyph_bits(w, h, seed):
    # 简单的"字母形"位图：边框 + 一道斜线（0 = 墨迹，1 = 透明）
    rows = []
    for y in range(h):
        row = []
        for x in range(w):
            ink = x == 0 or x == w - 1 or y == 0 or (x + seed) % w == (y * w) // h
            row.append(0 if ink else 1)
        bits = 0; out = bytearray()
        for i, b in enumerate(row):
            bits = (bits << 1) | b
            if i % 8 == 7: out.append(bits); bits = 0
        if w % 8: out.append(bits << (8 - w % 8))
        rows.append(bytes(out))
    return b"".join(rows)

def mask(x, y, w_pt, h_pt, seed, pw=8, ph=10):
    data = glyph_bits(pw, ph, seed)
    return (b"q %.2f 0 0 %.2f %.2f %.2f cm BI /IM true /W %d /H %d /BPC 1 /F /AHx ID " % (w_pt, h_pt, x, y, pw, ph)) + data.hex().encode() + b"> EI Q\n"

def text(x, y, s, size=10, font=b"F1"):
    s = s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    return b"BT /%s %d Tf %.2f %.2f Td (%s) Tj ET\n" % (font, size, x, y, s.encode("latin-1"))

c = bytearray()
# 两行居中的标题（以正文为准居中）；右上角的页码把版心撑得比正文宽
c += text(560, 765, "631", 8)
c += text(167.2, 742, "A Physical-Flow-Based Approach to Allocating", 16)
c += text(156.0, 724, "Transmission Losses in a Transaction Framework", 16)
# 正文段落（左右两端对齐的三行），第二行中间是位图字形组成的行内公式，第三行中间是 Type 3 矢量字形
c += text(72, 700, "The proposed allocation scheme assigns the total system losses to every")
c += text(72, 688, "transaction in proportion to its flow; the loss of transaction")
for i, x in enumerate([340, 347, 354]): c += mask(x, 686, 6, 8, i)
c += text(366, 688, ", which is computed below, is always")
c += text(72, 676, "nonnegative and the sum over all transactions equals the total loss")
c += b"BT /F2 10 Tf 368 676 Td <01> Tj ET\n"
c += text(376, 676, "of the system as a whole.")
# 独立成行的公式（位图字形 + 分数线）与右端编号
for i, x in enumerate([250, 258, 266, 280, 288, 296, 304]): c += mask(x, 640, 7, 9, i + 3)
c += b"0 g 312 643 m 340 643 l 340 643.6 l 312 643.6 l f\n"
for i, x in enumerate([316, 324]): c += mask(x, 646, 6, 7, i + 5)
for i, x in enumerate([316, 324]): c += mask(x, 634, 6, 7, i + 7)
c += text(520, 640, "(1)")
c += text(72, 610, "After the equation the text continues as a normal paragraph of body")
c += text(72, 598, "text so that the layout has more than one paragraph on the page.")
# 参考文献里画成短横线的"同上作者"破折号
c += text(72, 560, "[1] R. Nadira et al., Bulk Transmission System Loss Analysis, 1993.", 8)
c += text(72, 550, "[2]", 8)
c += b"0 g 86 552.6 m 102 552.6 l 102 553.0 l 86 553.0 l f\n"
c += text(104, 550, ", Bulk Transmission System Loss Analysis, EPRI Report, 1990.", 8)

objs = []
def add(b): objs.append(b); return len(objs)
font1 = add(b"<< /Type /Font /Subtype /Type1 /BaseFont /Times-Roman >>")
proc = b"600 0 0 0 500 700 d1 0 0 500 700 re f"
cp = add(b"<< /Length %d >>\nstream\n" % len(proc) + proc + b"\nendstream")
font2 = add(b"<< /Type /Font /Subtype /Type3 /FontBBox [0 0 600 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /g1 %d 0 R >> /Encoding << /Type /Encoding /Differences [1 /g1] >> /FirstChar 1 /LastChar 1 /Widths [600] /Resources << >> >>" % cp)
content = add(b"<< /Length %d >>\nstream\n" % len(c) + bytes(c) + b"endstream")
page = add(b"<< /Type /Page /Parent 6 0 R /MediaBox [0 0 612 792] /Contents %d 0 R /Resources << /Font << /F1 %d 0 R /F2 %d 0 R >> >> >>" % (content, font1, font2))
pages = add(b"<< /Type /Pages /Kids [%d 0 R] /Count 1 >>" % page)
assert pages == 6
cat = add(b"<< /Type /Catalog /Pages 6 0 R >>")
out = bytearray(b"%PDF-1.4\n")
offs = []
for i, o in enumerate(objs, 1):
    offs.append(len(out)); out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
xref = len(out)
out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
for o in offs: out += b"%010d 00000 n \n" % o
out += b"trailer\n<< /Size %d /Root %d 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, cat, xref)
open(sys.argv[1], "wb").write(out)
