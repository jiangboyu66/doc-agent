"""生成测试夹具 fontmatrix.pdf：字形宽度以 1/2048 em 为单位的字体（FontMatrix 0.00048828125）。
Elsevier 等出版社排版的 Charis SIL 就是这样：按常见的 1/1000 去读字宽，会把每个字算宽一倍，
进而把字符间距压成每字 −3 磅、Word 里文字挤成一团。这里用一个 Type 3 字体复现同样的字体矩阵。"""
import sys

UPM = 2048
GLYPHS = {"a": 1100, "b": 1150, "c": 950, "d": 1150, "e": 1000, "n": 1150, "o": 1100, "s": 850, "t": 700, "space": 520}
CODES = {"space": 32, **{ch: ord(ch) for ch in GLYPHS if ch != "space"}}

objs = []
def add(b): objs.append(b); return len(objs)
procs = {}
for name, w in GLYPHS.items():
    body = b"%d 0 0 0 %d 1400 d1" % (w, w) + (b"" if name == "space" else b" 60 0 %d 1000 re f" % (w - 120))
    procs[name] = add(b"<< /Length %d >>\nstream\n" % len(body) + body + b"\nendstream")
first, last = 32, max(CODES.values())
widths = [GLYPHS[n] if (n := next((k for k, v in CODES.items() if v == c), None)) else 0 for c in range(first, last + 1)]
diffs = b" ".join(b"%d /%s" % (CODES[n], n.encode()) for n in GLYPHS)
cmap = ("/CIDInit /ProcSet findresource begin 12 dict begin begincmap /CMapName /T3 def 1 begincodespacerange <00> <FF> endcodespacerange "
        + "%d beginbfchar " % len(CODES) + " ".join("<%02X> <%04X>" % (c, 32 if n == "space" else ord(n)) for n, c in CODES.items())
        + " endbfchar endcmap CMapName currentdict /CMap defineresource pop end end").encode()
tu = add(b"<< /Length %d >>\nstream\n" % len(cmap) + cmap + b"\nendstream")
charprocs = b" ".join(b"/%s %d 0 R" % (n.encode(), i) for n, i in procs.items())
font = add(b"<< /Type /Font /Subtype /Type3 /FontBBox [0 0 1200 1400] /FontMatrix [0.00048828125 0 0 0.00048828125 0 0] "
           b"/CharProcs << %s >> /Encoding << /Type /Encoding /Differences [%s] >> /FirstChar %d /LastChar %d /Widths [%s] /ToUnicode %d 0 R /Resources << >> >>"
           % (charprocs, diffs, first, last, b" ".join(b"%d" % w for w in widths), tu))
body = bytearray()
for k, line in enumerate(["coded notes on a dense dataset", "sent to the stations at dawn", "cats and dogs and ants"]):
    body += b"BT /F1 10 Tf 72 %d Td (%s) Tj ET\n" % (700 - 14 * k, line.encode())
content = add(b"<< /Length %d >>\nstream\n" % len(body) + bytes(body) + b"endstream")
pages_id = len(objs) + 2
page = add(b"<< /Type /Page /Parent %d 0 R /MediaBox [0 0 612 792] /Contents %d 0 R /Resources << /Font << /F1 %d 0 R >> >> >>" % (pages_id, content, font))
pages = add(b"<< /Type /Pages /Kids [%d 0 R] /Count 1 >>" % page)
assert pages == pages_id
cat = add(b"<< /Type /Catalog /Pages %d 0 R >>" % pages)
out = bytearray(b"%PDF-1.4\n")
offs = []
for i, o in enumerate(objs, 1):
    offs.append(len(out)); out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
xref = len(out)
out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
for o in offs: out += b"%010d 00000 n \n" % o
out += b"trailer\n<< /Size %d /Root %d 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, cat, xref)
open(sys.argv[1] if len(sys.argv) > 1 else "fontmatrix.pdf", "wb").write(out)
