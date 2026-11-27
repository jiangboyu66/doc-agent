"""生成测试夹具 rtl-list.pdf：希伯来文（从右往左）的正文和编号列表，最后一段英文。
列表编号像 WeasyPrint 那样单独画在行的右端，显示顺序是 ".1"（句点在数字左边）。
以前按从左往右排：编号排到了句末、显示成 ".1"，可编辑排版下第 1、2 条列表项并成一段。
reportlab 不做双向排版：希伯来文按显示顺序（逻辑顺序倒过来）画出，右对齐。"""
import sys
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

out = sys.argv[1] if len(sys.argv) > 1 else "rtl-list.pdf"
pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DejaVu-Bold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))
W, H = 595, 842
R = 523  # 右边界
c = canvas.Canvas(out, pagesize=(W, H))


def y(v):
    return H - v


def rtl(base, text, font="DejaVu", size=11, right=R):
    """一行希伯来文：显示顺序 = 逻辑顺序倒过来，右端对齐到 right"""
    c.setFont(font, size)
    c.drawRightString(right, y(base), text[::-1])


rtl(80, "סיכום תחבורה ציבורית", "DejaVu-Bold", 16)
# 正文段落：按版心宽度（72–523）逐词折行，像真正的段落那样排满
text = ("מערכת התחבורה הציבורית בעיר עברה שינוי משמעותי בשנה האחרונה עם קווים חדשים ואוטובוסים חשמליים "
        "ומספר הנוסעים עלה לעומת השנה הקודמת והזמן הממוצע לנסיעה ירד בצורה ניכרת בכל שעות היום ובמיוחד "
        "בשעות העומס של הבוקר והערב.")
lines, cur = [], ""
for w in text.split():
    t = (cur + " " + w).strip()
    if pdfmetrics.stringWidth(t, "DejaVu", 11) > R - 72 and cur:
        lines.append(cur)
        cur = w
    else:
        cur = t
lines.append(cur)
for k, s in enumerate(lines):
    rtl(115 + 17 * k, s)
items = ["להרחיב את רשת הקווים בשכונות החדשות.", "להחליף את כל האוטובוסים הישנים בחשמליים.", "להוסיף נתיבי תחבורה ציבורית בכבישים הראשיים."]
for k, s in enumerate(items):
    base = 190 + 19 * k
    rtl(base, s, right=R - 18)
    # 编号：显示顺序 ".1"，画在行的右端
    c.setFont("DejaVu", 11)
    c.drawRightString(R, y(base), f".{k + 1}")
c.setFont("DejaVu", 11)
c.drawString(72, y(270), "English summary: the Hebrew paragraph and list above run right to left.")
c.showPage()
c.save()
