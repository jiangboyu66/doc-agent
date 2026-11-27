"""生成测试夹具 form-boxes.pdf：申请表。
一、带框的大表格里，"身份证号"那一格里画着一排 18 个字符方格，右侧合并的一格里画着照片框（里面写"照片"）；
方格和照片框的线不碰大表格的格线，各自成一张小表格。以前它们作为单独的表格排在大表格后面：
大表格里留着空格子，"照片"两个字出现两次，后面的内容被挤到第二页。
二、表格格子里的横线："人民币（大写）______ ¥______"是与文字同一基线的填写线，"备注"一格里是两条不挨着文字的书写线。
以前可编辑排版把它们当成正文的横线，排在整张表格的后面。"""
import sys
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

out = sys.argv[1] if len(sys.argv) > 1 else "form-boxes.pdf"
pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
F = "STSong-Light"
W, H = 595, 842
c = canvas.Canvas(out, pagesize=(W, H))


def y(v):
    return H - v


c.setFont(F, 16)
c.drawCentredString(W / 2, y(70), "职工住房公积金提取申请表")
c.setFont(F, 11)
c.drawString(57, y(110), "一、申请人基本信息")
# 大表格：4 行，列 57 | 127 | 297 | 367 | 457 | 538（最右一列合并 4 行放照片框）
xs = [57, 127, 297, 367, 457, 538]
ys = [120, 150, 180, 210, 240]
c.setLineWidth(0.6)
for k, yy in enumerate(ys):
    c.line(xs[0], y(yy), xs[4] if 0 < k < 4 else xs[5], y(yy))
for x in (xs[0], xs[1], xs[4], xs[5]):
    c.line(x, y(ys[0]), x, y(ys[-1]))
# 第 2、3 行（身份证号、工作单位）合并了中间三列：这两条竖线只画在第 1、4 行
for x in (xs[2], xs[3]):
    c.line(x, y(ys[0]), x, y(ys[1]))
    c.line(x, y(ys[3]), x, y(ys[4]))
c.setFont(F, 10)
labels = [("姓名", "性别"), ("身份证号", None), ("工作单位", None), ("联系电话", "个人账号")]
for r, (a, b) in enumerate(labels):
    c.drawCentredString((xs[0] + xs[1]) / 2, y(ys[r] + 19), a)
    if b:
        c.drawCentredString((xs[2] + xs[3]) / 2, y(ys[r] + 19), b)
# 身份证号一格里的 18 个字符方格：相邻的方格共用一条边，连成一排；整排不碰外面的格线
bx, bw = 134, 16
for k in range(18):
    c.rect(bx + k * bw, y(ys[1] + 24), bw, 18, stroke=1, fill=0)
# 照片框
c.setDash(2, 2)
c.rect(xs[4] + 10, y(ys[-1] - 10), xs[5] - xs[4] - 20, ys[-1] - ys[0] - 20, stroke=1, fill=0)
c.setDash()
c.setFont(F, 10.5)
c.drawCentredString((xs[4] + xs[5]) / 2, y((ys[0] + ys[-1]) / 2 + 4), "照片")
c.setFont(F, 11)
c.drawString(57, y(280), "二、提取金额")
# 第二张表：2 行 2 列
xs2, ys2 = [57, 127, 538], [290, 320, 370]
for yy in ys2:
    c.line(xs2[0], y(yy), xs2[-1], y(yy))
for x in xs2:
    c.line(x, y(ys2[0]), x, y(ys2[-1]))
c.setFont(F, 10)
c.drawCentredString((xs2[0] + xs2[1]) / 2, y(309), "提取金额")
c.drawCentredString((xs2[0] + xs2[1]) / 2, y(349), "备注")
c.setFont(F, 10.5)
c.drawString(134, y(309), "人民币（大写）")
c.drawString(368, y(309), "¥")
# 填写线：与文字同一基线（略低 1pt），从文字后面一直画到格子里
c.line(210, y(310), 360, y(310))
c.line(380, y(310), 520, y(310))
# 备注格里两条书写线（同一高度没有文字）
c.line(134, y(343), 530, y(343))
c.line(134, y(362), 530, y(362))
c.setFont(F, 11)
c.drawString(57, y(400), "三、承诺")
c.setFont(F, 10.5)
c.drawString(78, y(425), "本人承诺所提供的材料真实、有效，如有虚假，愿意承担相应的法律责任。")
c.drawString(57, y(465), "申请人签字：")
c.drawString(400, y(465), "日期：     年   月   日")
c.showPage()
c.save()
