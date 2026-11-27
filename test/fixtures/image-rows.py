"""生成测试夹具 image-rows.pdf：两排矢量图形（每排六个，从版心左边开始，各自渲染成一张图），每个图形下面一行图名。
以前一排图写成一段、每张图前一个制表符，第一个制表位正好在 0 处：Word、LibreOffice 都跳过当前位置上的制表位，
整排图依次错后一个位置，最后一张折到下一行，下面的图名跟着下移。"""
import sys
from reportlab.pdfgen import canvas

out = sys.argv[1] if len(sys.argv) > 1 else "image-rows.pdf"
W, H = 595, 842
c = canvas.Canvas(out, pagesize=(W, H))


def y(v):
    return H - v


colors = [(220, 60, 60), (60, 160, 60), (60, 90, 220), (230, 180, 30), (150, 60, 180), (30, 170, 170)]
c.setFont("Helvetica-Bold", 14)
c.drawString(36, y(60), "Create basic shapes")
names = [["Triangle", "Right triangle", "Diamond", "Pentagon", "Hexagon", "Left arrow"],
         ["Right arrow", "Up arrow", "Down arrow", "Four point star", "Five point star", "Six point star"]]
for row, top in enumerate([90, 190]):
    for k in range(6):
        x = 36 + 91 * k
        r, g, b = colors[(k + row) % 6]
        c.setFillColorRGB(r / 255, g / 255, b / 255)
        c.setStrokeColorRGB(0, 0, 0)
        path = c.beginPath()
        # 每个图形是一条带曲线的封闭路径（不是方框，不会被当成表格或色块）
        path.moveTo(x + 30, y(top))
        path.curveTo(x + 62, y(top), x + 62, y(top + 40), x + 60, y(top + 60))
        path.lineTo(x, y(top + 60))
        path.curveTo(x - 2, y(top + 40), x - 2, y(top), x + 30, y(top))
        path.close()
        c.drawPath(path, stroke=1, fill=1)
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica", 9)
        c.drawString(x, y(top + 74), names[row][k])
c.setFont("Helvetica", 10)
c.drawString(36, y(300), "Each shape above has its name printed under it, in the same column.")
c.showPage()
c.save()
