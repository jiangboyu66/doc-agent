"""生成测试夹具 signature.pdf：合同末页的签字区与勾选表单（版面仿照常见的英文合同签字页）。
  - 签字区：左右两组签字线，线下居中的标签 "(Signature)" / "(Printed Name)" / "(Title)" 左右并排；
  - 复选框：描边的小方框（四条边，10.5pt），右侧是选项文字；第一个框里画了叉（已勾选）；
  - 无框表格：表头 "Prior Work / Party(ies) / Relationship"，每列一条比表头文字长的填写线，下面 3 行编号 + 填写线；
  - 框住字母 "A" 的小方框（不是复选框）。
这一页里左右并排的短标签、表头会被误判成"双栏正文"（栏边界取到编号 "1." 的行尾，标签被挤出栏外）；
方框的上下边太短被丢掉、只剩两条竖线；表头下的填写线被当成文字的下划线。"""
import sys
from reportlab.pdfgen import canvas

W, H = 612, 792
out = sys.argv[1] if len(sys.argv) > 1 else "signature.pdf"
c = canvas.Canvas(out, pagesize=(W, H))
R, B = "Times-Roman", "Times-Bold"


def text(x, base, s, font=R, size=11):
    """文字：x 为左端，base 为基线（页面坐标，y 向下）"""
    c.setFont(font, size)
    c.drawString(x, H - base, s)


def center(cx, base, s, font=R, size=11):
    c.setFont(font, size)
    c.drawString(cx - c.stringWidth(s, font, size) / 2, H - base, s)


def rule(x0, x1, y, h=0.48):
    """填写线：细长的填充矩形（y 为上边）"""
    c.rect(x0, H - y - h, x1 - x0, h, stroke=0, fill=1)


def box(x, y, checked=False):
    """复选框：描边方框（y 为上边）；已勾选的画一个叉"""
    c.setLineWidth(0.72)
    c.rect(x, H - y - 10.6, 10.5, 10.6, stroke=1, fill=0)
    if checked:
        c.line(x + 2, H - y - 2, x + 8.5, H - y - 8.6)
        c.line(x + 2, H - y - 8.6, x + 8.5, H - y - 2)


text(72, 82.3, "This Agreement takes effect on the date signed by the Consultant below.")
text(77.3, 106.8, "Example Holdings Ltd.", B)
text(351.8, 106.8, "Consultant:", B)
for y, base, left, right in [(147.4, 157.9, "(Signature)", "(Signature)"), (185.8, 196.6, "(Printed Name)", "(Printed Name)"), (224.2, 235.0, "(Title)", "(Date Signed)")]:
    rule(72, 324, y)
    rule(346.6, 598.6, y)
    center(198, base, left)
    center(472.6, base, right)
rule(108, 504, 261.4, 0.9)
center(306, 284.9, "PRIOR WORKS", B)
text(72, 309.6, "1.", B)
text(108, 309.6, "Prior Works Disclosure", B)
text(108 + c.stringWidth("Prior Works Disclosure", B, 11), 309.6, ". Except as listed in Section 2 below, the following is a complete list")
text(72, 322.3, "of all prior works:")
box(109, 337.4, checked=True)
text(144, 346.8, "No prior works.")
box(109, 362.2)
text(144, 371.5, "See below:")
for y in (398.4, 411.6, 424.8):
    rule(149.5, 540, y)
box(109, 439.0)
text(144, 448.3, "Additional sheets attached.")
text(72, 472.8, "2.", B)
text(108, 472.8, "Due to a prior confidentiality obligation, I cannot complete the disclosure under Section 1 above")
text(72, 485.5, "with respect to the prior works generally listed below, the rights and duty of confidentiality with")
text(72, 498.2, "respect to which I owe to the following party(ies):")
center(176.3, 522.7, "Prior Work", B)
center(338.4, 522.7, "Party(ies)", B)
center(496.0, 522.7, "Relationship", B)
for k, y in enumerate((525.1, 538.3, 551.5, 564.7)):
    for x0, x1 in ((113, 239.5), (253, 423.8), (437.5, 554.4)):
        rule(x0, x1, y)
    if k < 3:
        text(95.3, 535.9 + 13.2 * k, f"{k + 1}.", B)
box(109, 578.6)
text(144, 588.0, "Additional sheets attached.")
# 框住字母的小方框（不是复选框）：四条边都要画出来
text(72, 620.0, "See Exhibit")
box(130, 610.6)
center(135.25, 620.0, "A")
c.setFont(R, 8)
c.drawRightString(540, H - 734.9, "Consulting Agreement")
c.drawRightString(540, H - 744.2, "Page 6")
c.showPage()
c.save()
