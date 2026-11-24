"""生成测试夹具 furniture-cn.pdf：6 页中文合同，页眉页脚与恰好排在页首的正文。
  - 页眉 "设备采购合同（第3版）"（每页相同，"3" 是版本号，不是页码）；页脚 "第 N 页 共 6 页"（居中）；
  - 第 2、3 页的第一行正文分别是单独成行的条款标题 "第4条"、"第5条"（条款号减页序号恰好相同）；
  - 第 4–6 页每页一份附件，第一行是 "附件1"、"附件2"、"附件3"（同样逐页加一），以 "（附件完）" 结束、下一份另起一页。
易错点：按"数字归一后文字相同、位于同一高度、数字随页序递增"判页眉页脚，条款标题、附件标题被当成页眉移出正文，
Word 的页眉只能保留其中一个；代表页恰好是第 3 页时，页眉里的 "第3版" 被写成 PAGE 域，各页显示成 "第1版""第5版"。"""
import sys
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

W, H = 595.3, 841.9
out = sys.argv[1] if len(sys.argv) > 1 else "furniture-cn.pdf"
pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
F = "STSong-Light"
LEFT, TOP, PITCH, SIZE = 90.0, 110.0, 22.6, 10.5
c = canvas.Canvas(out, pagesize=(W, H))
TEXT = [
    "甲方应当在本合同签订之日起十日内向乙方支付首期款项，乙方收到款项后应当出具合法有效的发票。",
    "乙方保证其提供的货物符合国家标准及本合同附件所列技术要求，且不存在任何权利瑕疵。",
    "任何一方未履行本合同约定义务的，应当向守约方支付违约金，违约金不足以弥补损失的，还应赔偿差额。",
    "双方因履行本合同发生争议的，应友好协商解决；协商不成的，任何一方均可向人民法院提起诉讼。",
    "本合同项下的通知应以书面形式送达至首页所列地址，一方变更地址的应提前五个工作日书面通知对方。",
]
PAGES = [
    [("设备采购合同", 16), ("第1条", 12), 0, 1, ("第2条", 12), 2, 3, ("第3条", 12), 4, 0, 1, 2, 3, 4, 0, 1, 2, 3, 4, 0, 1, 2, 3, 4, 0, 1, 2],
    [("第4条", 12), 3, 4, 0, 1, 2, 3, 4, 0, 1, 2, 3, 4, 0, 1, 2, 3, 4, 0, 1, 2, 3, 4, 0, 1, 2, 3, 4, 0],
    [("第5条", 12), 1, 2, 3, 4, 0, 1, 2, 3, 4, ("第6条", 12), 0, 1, 2, 3, 4, 0, 1, 2, 3, 4, 0, 1, 2],
    [("附件1", 12), ("报价单", 12), 0, 1, 2, 3, 4, 0, 1, 2, ("（附件完）", SIZE)],
    [("附件2", 12), ("技术规格", 12), 2, 3, 4, 0, 1, 2, ("（附件完）", SIZE)],
    [("附件3", 12), ("授权委托书", 12), 4, 0, 1, 2, 3, ("（附件完）", SIZE)],
]
for p, items in enumerate(PAGES, 1):
    c.setFont(F, 9)
    c.drawString(LEFT, H - 50, "设备采购合同（第3版）")
    footer = f"第 {p} 页 共 {len(PAGES)} 页"
    c.drawString((W - pdfmetrics.stringWidth(footer, F, 9)) / 2, H - 805, footer)
    y = TOP
    for it in items:
        text, size = (it if isinstance(it, tuple) else (TEXT[it], SIZE))
        c.setFont(F, size)
        c.drawString(LEFT, H - y, text)
        y += PITCH
    c.showPage()
c.save()
