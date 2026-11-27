"""生成测试夹具 italic-rotated.pdf：中文专利扉页，带合成斜体和一段旋转的边注，内容为虚构。
  - IPC 分类号 "F16H 25/22"、"G01L 5/00" 和一行中文，按浏览器给没有斜体字形的字体合成斜体的方式绘制：
    倾斜的文字矩阵 [1 0 0.25 1]，不旋转；
  - 左边距有一段旋转 90° 的边注（"合成测试文件 仅供测试"），不在任何图里。
易错点：1.1.2 把倾斜的文字当成旋转文字丢掉，文字完整性仍报 100%（2026-11-27 LTB-05：扉页的 IPC 分类号消失）。
标准答案：倾斜的文字是正文、斜体；旋转的边注不转换，但计入文字完整性的缺失，并在说明里点名。"""
import sys
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

W, H = 595.0, 842.0
out = sys.argv[1] if len(sys.argv) > 1 else "italic-rotated.pdf"
pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
CN = "STSong-Light"
c = canvas.Canvas(out, pagesize=(W, H))


def sheared(x: float, y: float, text: str, font: str, size: float) -> None:
    t = c.beginText()
    t.setFont(font, size)
    t.setTextTransform(1, 0, 0.25, 1, x, y)
    t.textOut(text)
    c.drawText(t)


c.setFont(CN, 16)
c.drawString(72, 780, "（12）发明专利申请")
c.setFont(CN, 10.5)
c.drawString(72, 740, "（21）申请号 202610999999.9")
c.drawString(72, 722, "（22）申请日 2026.02.06")
c.drawString(330, 740, "（51）Int.Cl.")
sheared(360, 722, "F16H 25/22 (2006.01)", "Times-Roman", 10.5)
sheared(360, 704, "G01L 5/00 (2006.01)", "Times-Roman", 10.5)
c.drawString(72, 660, "（54）发明名称")
sheared(150, 660, "一种合成测试用的丝杠传动装置", CN, 10.5)
c.drawString(72, 620, "（57）摘要")
c.drawString(72, 600, "本发明公开了一种丝杠传动装置，包括丝杠、螺母和用于检测轴向力的传感器。本文件为合成测试文件。")

# 左边距旋转的边注：不转换，计入缺失
c.saveState()
c.translate(40, 300)
c.rotate(90)
c.setFont(CN, 9)
c.drawString(0, 0, "合成测试文件 仅供测试")
c.restoreState()
c.showPage()
c.save()
