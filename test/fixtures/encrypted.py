"""生成测试夹具 encrypted.pdf：设了打开密码（用户密码 doc-agent）的一页英文文字 PDF。
用途：pdf2word 发布包的冒烟测试——没有密码时 pdf.js 抛出 PasswordException，命令行应返回 code=encrypted，而不是笼统的转换失败。"""
import sys
from reportlab.pdfgen import canvas

W, H = 612, 792
out = sys.argv[1] if len(sys.argv) > 1 else "encrypted.pdf"
c = canvas.Canvas(out, pagesize=(W, H), invariant=1, encrypt="doc-agent")
c.setFont("Times-Roman", 12)
c.drawString(72, H - 100, "This document is protected with an open password.")
c.drawString(72, H - 120, "It must not be converted without the password.")
c.showPage()
c.save()
