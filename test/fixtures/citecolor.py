"""生成测试夹具 citecolor.pdf：期刊正文里蓝色的引用编号（IEEE Access 等的链接引用）。
引用编号与后面的黑色正文在同一个文字对象里、字体相同，只是填充颜色不同——pdf.js 会把它们合成一个文字项。
行首是蓝色引用的那一行（"[23]. While the transformer method is good at capturing …"），
以前整行取第一个字的颜色、被染成蓝色；行中的引用反而丢了颜色。"""
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color

out = sys.argv[1] if len(sys.argv) > 1 else "citecolor.pdf"
c = canvas.Canvas(out, pagesize=(612, 792))
blue, black = Color(0, 0, 1), Color(0, 0, 0)


def line(y, parts, x=72):
    t = c.beginText(x, 792 - y)
    t.setFont("Times-Roman", 10)
    for col, s in parts:
        t.setFillColor(col)
        t.textOut(s)
    c.drawText(t)


line(100, [(black, "Work in "), (blue, "[18]"), (black, " employs a transformer-based approach as a fully end-to-end object detector,")])
line(112, [(black, "delving into the connections between objects and the overall image context for object detection.")])
line(124, [(black, "There are several transformer-based models designed for a specific task "), (blue, "[20], [21], [22],")])
line(136, [(blue, "[23]"), (black, ". While the transformer method is good at capturing global context, it struggles to grasp")])
line(148, [(black, "fine-grained details, especially for medical images. Efforts have been made to integrate CNN-")])
line(160, [(black, "and Transformer-based models into each other "), (blue, "[24], [25]"), (black, ". In particular, TransUNet "), (blue, "[29]")])
line(172, [(blue, "[27], [28]"), (black, ". In particular, TransUNet and TransFuse are representative examples of combining")])
line(184, [(black, "the transformer and UNet for medical image segmentation.")])
c.showPage()
c.save()
