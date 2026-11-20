"""生成测试夹具 contract-cn.pdf：中文律所合同草稿的 3 页（版面仿照 Word 导出的分配协议）。
  - 五号字（10.5pt）、1.5 倍行距：基线间距 22.6pt，是字号的 2.15 倍；条款之间段后 6pt；
  - 多级编号、悬挂缩进："1.1⇥正文…" 编号在 90pt，正文与续行在 111.3pt；"(a)⇥…" 编号在 111.3pt，正文在 132.6pt；
    编号 "10.12" 比悬挂缩进宽，首行正文被推到下一个制表位，续行仍从 111.3pt 起排；
  - 方括号占位符 [ ]、[30 日]、[●] 以黄色突出显示：每个文字片段背后一个紧贴的填充矩形（英文 12pt 高、中文 15pt 高）；
  - 第 1 条款 1.3 从第 1 页末跨到第 2 页；
  - 页眉 "分配协议"（右）；页脚两行：左 "保密文件 · 律师工作草稿"、右 "第 N 页 / 共 3 页"，第一行页脚距正文末行只有约 6pt。
易错点：按"行距 ≤ 1.75 倍字号"合并段落时每一行都成了一段；"1.1"不在列表符号里、续行对不上编号所在的左边界；
高亮矩形被当成矢量图 / 行内小图标，占位符连同里面的文字被截成图片；页脚上方没有明显空白，两行页脚留在正文里、截断跨页的条款。
标准答案（每个条款的全文）写在 contract-cn.truth.json。"""
import json, os, sys
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

W, H = 595.3, 841.9
out = sys.argv[1] if len(sys.argv) > 1 else "contract-cn.pdf"
pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
CJK, LAT = "STSong-Light", "Times-Roman"
SIZE, PITCH, AFTER = 10.5, 22.6, 6.0
LEFT, RIGHT = 90.0, 505.3
c = canvas.Canvas(out, pagesize=(W, H))

PH = {"[ ]", "[30 日]", "[●]"}


def runs_of(text):
    """按字体切分：中文（含全角标点）用宋体，其余用 Times；占位符单独成段（突出显示）"""
    out, i = [], 0
    while i < len(text):
        hit = next((p for p in PH if text.startswith(p, i)), None)
        if hit:
            for ch_run in split_font(hit): out.append((ch_run[0], ch_run[1], True))
            i += len(hit); continue
        j = i
        while j < len(text) and not any(text.startswith(p, j) for p in PH): j += 1
        for r in split_font(text[i:j]): out.append((r[0], r[1], False))
        i = j
    return out


def split_font(s):
    res = []
    for ch in s:
        f = CJK if ord(ch) > 0x2E80 or ch in "●" else LAT
        if res and res[-1][1] == f: res[-1][0] += ch
        else: res.append([ch, f])
    return [(t, f) for t, f in res]


def width(t, f):
    return pdfmetrics.stringWidth(t, f, SIZE)


def wrap(text, first_w, rest_w):
    """按宽度折行（逐字），返回每行的文字"""
    lines, cur, avail = [], "", first_w
    for ch in text:
        if cur and sum(width(t, f) for t, f, _ in runs_of(cur + ch)) > avail:
            lines.append(cur); cur, avail = "", rest_w
        cur += ch
    if cur: lines.append(cur)
    return lines


class Page:
    def __init__(self):
        self.n = 0
        self.y = 0.0

    def new(self):
        if self.n: footer(self.n); c.showPage()
        self.n += 1
        self.y = 84.1
        c.setFont(CJK, 9); c.drawString(469.6, H - 52.9, "分配协议")


pg = Page()


def footer(n):
    c.setFont(CJK, 9); c.drawString(LEFT, H - 777.9, "保密文件 · 律师工作草稿")
    s = f"第 {n} 页 / 共 3 页"
    c.drawString(RIGHT - sum(pdfmetrics.stringWidth(ch, CJK if ord(ch) > 0x2E80 else CJK, 9) for ch in s), H - 790.8, s)


def draw_line(x, base, text, justify_to=None):
    rs = runs_of(text)
    nat = sum(width(t, f) for t, f, _ in rs)
    n = len(text)
    cs = (justify_to - x - nat) / max(1, n - 1) if justify_to else 0.0
    for t, f, hl in rs:
        w = width(t, f) + cs * len(t)
        if hl:
            h = 15.0 if f == CJK else 12.0
            top = base - (11.5 if f == CJK else 9.4)
            c.setFillColorRGB(1, 1, 0); c.rect(x, H - top - h, w, h, stroke=0, fill=1); c.setFillColorRGB(0, 0, 0)
        to = c.beginText(x, H - base); to.setFont(f, SIZE); to.setCharSpace(cs); to.textOut(t); c.drawText(to)
        x += w


def clause(label, text, label_x, text_x, cont_x=None):
    """悬挂缩进的条款：编号在 label_x，首行正文在 text_x，续行在 cont_x（默认同 text_x）"""
    cont_x = cont_x if cont_x is not None else text_x
    lines = wrap(text, RIGHT - text_x, RIGHT - cont_x)
    for k, ln in enumerate(lines):
        if pg.y > 765: pg.new()
        if k == 0:
            c.setFont(LAT, SIZE); c.drawString(label_x, H - pg.y, label)
        x = text_x if k == 0 else cont_x
        draw_line(x, pg.y, ln, RIGHT if k < len(lines) - 1 else None)
        pg.y += PITCH
    pg.y += AFTER
    truth.append(label + text)


truth = []
S = [
    "对于前款所述事项，双方应本着诚实信用的原则友好协商解决；协商不成的，任何一方均可依照本协议第二十条的约定解决争议。",
    "分销商应建立完善的客户档案和销售台账，并于每季度结束后十个工作日内向供应商提交上一季度的销售报告、库存报告及市场分析报告。",
    "未经对方事先书面同意，任何一方不得将其在本协议项下的全部或部分权利、义务转让给任何第三方，但向其关联方转让的除外。",
    "甲方对协议产品的定价享有最终决定权，但应提前不少于三十日将价格调整方案以书面形式通知乙方。",
]
pg.new()
c.setFont(CJK, 10.5); c.drawString(LEFT, H - pg.y, "第 1 条　定义与解释"); pg.y += PITCH + AFTER
clause("1.1", S[0] + "乙方应于本协议生效之日起[30 日]内完成上述事项。" + S[1], LEFT, 111.3)
clause("1.2", S[2] + "供应商应在收到订单后[ ]内向对方支付人民币[●]元（含税）。", LEFT, 111.3)
clause("(a)", S[3] + S[0], 111.3, 132.6)
clause("(b)", "双方同意于[ ]内书面答复对方。" + S[1], 111.3, 132.6)
for k in range(3):
    clause(f"1.{3 + k}", S[k % 4] + S[(k + 1) % 4] + S[(k + 2) % 4] + "双方确认，上述约定系在平等自愿的基础上协商一致作出的。", LEFT, 111.3)
# 编号比悬挂缩进宽：首行正文推到下一个制表位（90 + 2 × 21.3）
clause("10.12", S[1] + "乙方应于本协议生效之日起[30 日]内向对方交付相关文件。" + S[3], LEFT, LEFT + 42.6, 111.3)
for k in range(12):
    clause(f"2.{k + 1}", S[(k + 1) % 4] + S[(k + 3) % 4] + ("供应商应在收到订单后[ ]内书面答复对方。" if k % 2 else ""), LEFT, 111.3)
footer(pg.n)
c.save()
assert pg.n == 3, pg.n
json.dump(truth, open(os.path.splitext(out)[0] + ".truth.json", "w"), ensure_ascii=False, indent=0)
