"""生成测试夹具 cn-cidfont.pdf：只用未嵌入的中文 CID 字体（reportlab 的 STSong-Light，编码 UniGB-UCS2-H）排版的一页协议。
页面上每个字都是这一种字体：没有嵌入字形，也没有 ToUnicode 表，字符编码到 Unicode 全靠 pdf.js 自带的 CMap 数据
（cmaps/UniGB-UCS2-H.bcmap、Adobe-GB1-UCS2.bcmap）。
用途：pdf2word 发布包的冒烟测试——包里漏了 cmaps/ 时，这份 PDF 提取出的文字是 0 个（不是"少几个"），完整性校验却仍是 100%。"""
import sys
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

W, H = 595.3, 841.9
out = sys.argv[1] if len(sys.argv) > 1 else "cn-cidfont.pdf"
pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
FONT = "STSong-Light"

c = canvas.Canvas(out, pagesize=(W, H), invariant=1)
c.setTitle("保密协议")

c.setFont(FONT, 16)
title = "保密协议"
c.drawString((W - c.stringWidth(title, FONT, 16)) / 2, H - 100, title)

clauses = [
    "第一条　定义",
    "本协议所称保密信息，是指披露方以任何形式向接收方提供的、尚未公开的技术信息与经营信息。",
    "第二条　保密义务",
    "接收方应当对保密信息严格保密，未经披露方书面同意，不得向任何第三方披露。",
    "接收方仅得为履行本协议之目的使用保密信息，并应当采取不低于保护自身同类信息的措施。",
    "第三条　期限",
    "本协议自双方签字盖章之日起生效，保密义务在本协议终止后继续有效。",
    "第四条　争议解决",
    "因本协议引起的争议，双方应当友好协商解决；协商不成的，提交合同签订地人民法院诉讼解决。",
]
y = H - 150
c.setFont(FONT, 10.5)
for s in clauses:
    c.drawString(90, y, s)
    y -= 26
c.showPage()
c.save()
