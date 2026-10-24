"""生成测试夹具 bios.pdf：模仿 IEEE 期刊最后一页的作者简介。
两栏；每位作者一张方形照片，放大后用取景框裁成 72 × 90 pt 的竖幅（PDF 里的图片外接框比实际显示的大），
正文紧挨着照片右侧排列，照片下方的行占满整栏。左栏两位作者的照片上下间距很小（图片外接框会互相接近）。
照片是程序生成的渐变人像，不含真实人物。"""
import io, sys
from PIL import Image, ImageDraw
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

W, H = 576, 783
out = sys.argv[1] if len(sys.argv) > 1 else "bios.pdf"

def photo(seed):
    im = Image.new("RGB", (400, 400))
    d = ImageDraw.Draw(im)
    for y in range(400):
        d.line([(0, y), (400, y)], fill=(60 + seed * 30 + y // 8, 90 + y // 6, 140 - seed * 20))
    d.ellipse((120, 70, 280, 260), fill=(210, 170, 140))
    d.rectangle((90, 270, 310, 400), fill=(40 + seed * 25, 40, 70))
    return ImageReader(im)

BIOS = [
    ("L", 65, "MUHAMMAD IMRAN", "received the bachelor's degree from Mehran University, Jamshoro, the M.S. degree from NUST, Islamabad, and the Ph.D. degree from Florida State University, USA. He is currently a Postdoctoral Associate of artificial intelligence with the College of Medicine, University of Florida. He is working on tissue image analysis, particularly for the diagnosis of cancer. His research interests include computer vision and deep learning for medical image analysis."),
    ("L", 202, "BUSHRA HAQ", "(Student Member, IEEE) received the M.S. degree from the Balochistan University of Information Technology, Engineering, and Management Sciences (BUITEMS), where she is currently pursuing the Ph.D. degree. Since 2014, she has been affiliated with BUITEMS as a Faculty Member with FICT. Her research interests include data science and remote sensing."),
    ("L", 326, "ERSIN ELBASI", "received the M.Sc. degree in computer science from Syracuse University, and the M.Phil. and Ph.D. degrees in computer science from the Graduate Center, The City University of New York. He is currently with American University of the Middle East. His research interests include multimedia security, event mining in video sequences, and medical image processing."),
    ("R", 65, "AHMET E. TOPCU", "(Member, IEEE) received the B.Sc. degree from the Department of Electrical and Electronics Engineering, Middle East Technical University, Ankara, Turkey, in 1997; the M.Sc. degree in computer engineering from Syracuse University, Syracuse, NY, USA, in 2001; and the Ph.D. degree in computer science from Indiana University, Bloomington, IN, USA, in 2010. He is currently an Associate Professor with the College of Engineering and Technology, American University of the Middle East, Kuwait. His current research interests include distributed systems, machine learning, cloud computing, big data, the IoT, and blockchain."),
    ("R", 318, "WEI SHAO", "(Member, IEEE) is currently an Assistant Professor and the Assistant Director of AI imaging with the Intelligent Critical Care Center, University of Florida. He is the Principal Investigator of the UF Medical Imaging Research for Translational Healthcare with Artificial Intelligence laboratory, where he leads a team of researchers in developing and applying deep learning and medical image analysis methods for early disease diagnosis and clinical decision support."),
]
COL = {"L": (36, 242), "R": (297, 243)}
FS, LEAD = 9.5, 9.55
PHOTO_W, PHOTO_H, GAP = 72, 90, 10

c = canvas.Canvas(out, pagesize=(W, H))
c.setFont("Helvetica", 7); c.drawString(36, H - 31, "M. Imran et al.: Transformer-Based Hierarchical Model for NSCLC Detection and Classification")
c.setLineWidth(0.4); c.line(36, H - 37, 538, H - 37)
c.drawString(36, 38, "VOLUME 12, 2024"); c.drawRightString(540, 38, "145933")

def wrap(words, width, font, size):
    lines, cur = [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if c.stringWidth(t, font, size) <= width: cur = t
        else: lines.append(cur); cur = w
    if cur: lines.append(cur)
    return lines

for i, (side, top, name, body) in enumerate(BIOS):
    x0, cw = COL[side]
    # 照片：方形原图放大到 100 × 100，取景框 72 × 90 居中裁剪
    c.saveState()
    p = c.beginPath(); p.rect(x0, H - top - PHOTO_H, PHOTO_W, PHOTO_H)
    c.clipPath(p, stroke=0, fill=0)
    c.drawImage(photo(i), x0 - 14, H - top - PHOTO_H - 5, 100, 100)
    c.restoreState()
    # 正文：照片旁每行缩进，照片下方占满整栏
    c.setFont("Helvetica-Bold", 8.5); nw = c.stringWidth(name + " ", "Helvetica-Bold", 8.5)
    words = body.split(" ")
    y = H - top - FS
    x_beside = x0 + PHOTO_W + GAP
    first = True
    while words:
        beside = y >= H - top - PHOTO_H - 6.7
        xs = x_beside if beside else x0
        width = (x0 + cw - xs) - (nw if first else 0)
        lines = wrap(words, width, "Times-Roman", FS)
        line = lines[0]
        n = len(line.split(" "))
        if first:
            c.setFont("Helvetica-Bold", 8.5); c.drawString(xs, y, name)
            c.setFont("Times-Roman", FS); c.drawString(xs + nw, y, line)
        else:
            c.setFont("Times-Roman", FS); c.drawString(xs, y, line)
        first = False
        words = words[n:]
        y -= LEAD
c.showPage(); c.save()
