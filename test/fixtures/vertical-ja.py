"""生成测试夹具 vertical-ja.pdf：日文縦書き（竖排，从上往下、从右往左），两页 A5。
第 1 页：横排字体逐字摆放（LibreOffice、Chrome 打印的竖排就是这样），列里的 "PDF 2.0" 横躺着排（转 90°）；
第 2 页：竖排字体（WMode 1，UniJIS-UCS2-V，InDesign 等排的竖排是这样），一列一个文字项。
两页页底各有一个横排的页码。
以前逐字摆放的字按横排成行：同一高度上各列的字连成一"行"，正文顺序全乱，还被当成双栏，一页变两页；
竖排字体的一列挤在第一个字的位置上。"""
import sys
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

out = sys.argv[1] if len(sys.argv) > 1 else "vertical-ja.pdf"
W, H = 420, 595
pdfmetrics.registerFont(UnicodeCIDFont("HeiseiMin-W3"))
pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
_v = UnicodeCIDFont("HeiseiMin-W3", isVertical=True)
_v.name = _v.fontName = VMIN = "HeiseiMin-W3-V"
pdfmetrics.registerFont(_v)
c = canvas.Canvas(out, pagesize=(W, H))
TOP, BOTTOM, RIGHT = 50, 545, 380  # 版心：天头、地脚、第一列的右边
S, PITCH = 10.5, 18.9
PER = int((BOTTOM - TOP) // S)  # 一列的字数

P1 = [
    ("HeiseiKakuGo-W5", 16, "山あいの村の春", 0),
    ("HeiseiMin-W3", S, "三月の終わりになると、村を囲む山の斜面から少しずつ雪が消えていく。朝の空気はまだ冷たいが、日が昇ると屋根の雪がぽたぽたと音を立てて落ち、軒先には小さな水たまりがいくつもできる。", 1),
    ("HeiseiMin-W3", S, "子どもたちは学校の帰りに田んぼのあぜ道を通り、ふきのとうを探して歩く。夜は次の会議で{PDF 2.0}の話をする予定だと、父が言った。", 1),
]
P2 = [
    "四月に入ると、川の水かさが増えて流れが速くなる。雪どけの水は透きとおっていて、川底の石の色まではっきりと見える。",
    "祖父は毎年この時期に水車小屋の点検をし、古くなった板を一枚ずつ取り替える。桜が咲くのは平地より二週間ほど遅い。",
]


def page_number(n):
    c.setFont("Helvetica", 9)
    c.drawCentredString(W / 2, H - 572, str(n))


def glyph_page():
    """逐字摆放：每个字一个文字项，列从右往左，列内从上往下；{...} 里的拉丁字母横躺着排"""
    x = RIGHT
    for font, size, text, indent in P1:
        pitch = max(PITCH, size * 1.6)
        cx = x - size / 2
        k = indent  # 本列已排的字数（首行缩进一字）
        i = 0
        while i < len(text):
            if k >= PER:
                x -= pitch
                cx = x - size / 2
                k = 0
            if text[i] == "{":
                j = text.index("}", i)
                latin = text[i + 1:j]
                c.saveState()
                # 基线在列中线左边 0.35 字号，字头朝右，从上往下读
                c.translate(cx - 0.35 * size, H - (TOP + k * size))
                c.rotate(-90)
                c.setFont("Helvetica", size)
                c.drawString(0, 0, latin)
                c.restoreState()
                k += int(c.stringWidth(latin, "Helvetica", size) // size) + 1
                i = j + 1
                continue
            c.setFont(font, size)
            c.drawString(cx - size / 2, H - (TOP + k * size + 0.88 * size), text[i])
            k += 1
            i += 1
        x -= pitch
    page_number(1)


def vertical_font_page():
    """竖排字体：一列一个文字项（文字位置是列顶的中点）"""
    x = RIGHT
    for text in P2:
        first = True
        rest = text
        while rest:
            n = PER - (1 if first else 0)
            col, rest = rest[:n], rest[n:]
            c.setFont(VMIN, S)
            c.drawString(x - S / 2, H - (TOP + (S if first else 0)), col)
            first = False
            x -= PITCH
    page_number(2)


glyph_page()
c.showPage()
vertical_font_page()
c.save()
