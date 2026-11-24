"""生成测试夹具 bates.pdf：4 页英文文件，页眉 "Production Volume 1 — Confidential"，右下角是逐页递增的 Bates 号
ABC0001234–ABC0001237（不是本文的页码，也没有另印页码）。
易错点：Bates 号被当成页眉页脚后，Word 页脚只能写一个固定号码，其余各页显示的号码都是错的；
放回正文又会截断跨页的段落。应当不写进页脚，并在 notes 里说明。"""
import sys
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics

W, H = 612.0, 792.0
out = sys.argv[1] if len(sys.argv) > 1 else "bates.pdf"
c = canvas.Canvas(out, pagesize=(W, H))
LINE = "The Supplier shall deliver the Goods to the Delivery Point on the Delivery Date in accordance with the Specifications"
for p in range(4):
    c.setFont("Helvetica", 9)
    c.drawString(72, H - 40, "Production Volume 1 — Confidential")
    bates = f"ABC{1234 + p:07d}"
    c.drawString(W - 72 - pdfmetrics.stringWidth(bates, "Helvetica", 9), 36, bates)
    c.setFont("Times-Roman", 11)
    y = H - 90
    for k in range(34):
        c.drawString(72, y, f"{p + 1}.{k + 1} {LINE[: 60 + (k * 7 + p * 11) % 50]}.")
        y -= 18
    c.showPage()
c.save()
