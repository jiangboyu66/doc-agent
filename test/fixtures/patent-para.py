"""生成测试夹具 patent-para.pdf：日文专利说明书的版面（仿照国际公开小册子），内容为虚构。
  - 第 1 页：著录项目两栏，右栏伸到 555pt，把全文的版心撑得比正文宽；
  - 第 2 页：说明书正文。段落编号 "[0006]" 在左边距（72pt），首行正文空一字从 126pt 起排，续行从 114pt 起排，
    行尾在 522pt；版心左右（74–555pt）两侧的留白与正文满行左右的留白相近。
易错点：续行落在编号和首行正文之间，既不对齐编号也不对齐首行正文，每段的首行单独成段；满行左右留白相近，
余下的行被当成居中。标准答案：每段一段、左对齐（或两端对齐），不居中。"""
import sys
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

W, H = 595.0, 842.0
out = sys.argv[1] if len(sys.argv) > 1 else "patent-para.pdf"
pdfmetrics.registerFont(UnicodeCIDFont("HeiseiMin-W3"))
F, SIZE, PITCH = "HeiseiMin-W3", 11.5, 24.0
BODY, FIRST, LABEL, RIGHT = 114.0, 126.0, 72.0, 522.0
c = canvas.Canvas(out, pagesize=(W, H))

# 第 1 页：两栏著录项目，右栏伸到 555pt
c.setFont(F, 10)
left = ["（５１）国際特許分類：", "Ｇ０６Ｔ　５／５０", "（２１）国際出願番号：", "（２２）国際出願日：", "（２５）国際出願の言語：　日本語"]
right = ["（７１）出願人：架空映像株式会社", "〒１００－０００１　東京都架空区一丁目１番１号", "（７２）発明者：山田　一郎", "（７４）代理人：弁理士　架空　次郎",
         "（８１）指定国：ＡＥ，ＡＧ，ＡＬ，ＡＭ，ＡＯ，ＡＴ，ＡＵ"]
for i, t in enumerate(left):
    c.drawString(74, 760 - 18 * i, t)
for i, t in enumerate(right):
    c.drawRightString(555, 760 - 18 * i, t)
c.setFont(F, 12)
c.drawString(74, 640, "（５４）発明の名称：映像処理装置、映像処理方法、及び、プログラム")
c.showPage()

# 第 2 页：说明书
PARAS = [
    ("[0006]", "本技術の一側面の映像処理装置は、異なる複数の視点位置において時系列で撮像された被写体を含む入力画像フレームに含まれる注目被写体に対する視点位置の移動の範囲を示す視点位置情報、及び、前記入力画像フレームの時間的な範囲を示す時間情報を取得する第１の取得部を備える映像処理装置である。"),
    ("[0007]", "本技術の一側面の映像処理方法及びプログラムは、上述した本技術の一側面の映像処理装置に対応する映像処理方法及びプログラムである。前記注目被写体の位置に応じて前記入力画像フレームを処理し、出力画像フレームを出力する。"),
    ("[0008]", "本技術の一側面によれば、被写体が動く場合においても、当該被写体を中心とした映像を提供することができる。なお、ここに記載された効果は必ずしも限定されるものではなく、本開示中に記載されたいずれかの効果であってもよい。"),
]
c.setFont(F, SIZE)
y = 760.0
per = int((RIGHT - BODY) // SIZE)  # 一行的字数（满行）
for label, text in PARAS:
    c.drawString(LABEL, y, label)
    first = int((RIGHT - FIRST) // SIZE)
    chunks = [text[:first]] + [text[first + k:first + k + per] for k in range(0, len(text) - first, per)]
    for i, chunk in enumerate(chunks):
        x = FIRST if i == 0 else BODY
        if i < len(chunks) - 1:
            # 两端对齐：满行的字均匀铺满到 522pt
            gap = (RIGHT - x - len(chunk) * SIZE) / max(1, len(chunk) - 1)
            for k, ch in enumerate(chunk):
                c.drawString(x + k * (SIZE + gap), y, ch)
        else:
            c.drawString(x, y, chunk)
        y -= PITCH
    y -= 4
c.showPage()
c.save()
