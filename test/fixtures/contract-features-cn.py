"""生成测试夹具 contract-features-cn.doc：中文合同（旧版 Word 97-2003 格式），覆盖 .doc 转换要处理的主要结构。
  - Word 自动编号：第一条 / 1.1 / (a)（中文计数、法律编号、小写字母），悬挂缩进；带圈数字项目符号
  - 表格：表头底纹、横向合并、纵向合并、边框
  - 行内图片、浮动图片（盖章）、脚注、批注、修订（插入 / 删除）、黄色高亮占位符
  - 页眉（带下边框）、页脚（第 PAGE 页 共 NUMPAGES 页）、第二节横向页面
本脚本生成 .docx，再用 LibreOffice 另存为 .doc：
  python contract-features-cn.py /tmp/features.docx
  soffice --headless --convert-to doc:"MS Word 97" --outdir /tmp /tmp/features.docx  # → /tmp/features.doc
"""
import io, sys, copy
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from docx.shared import Pt, Cm, Inches, RGBColor
from PIL import Image, ImageDraw

OUT = sys.argv[1]
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def png(w, h, draw):
    im = Image.new("RGBA", (w, h), (255, 255, 255, 0))
    draw(ImageDraw.Draw(im))
    b = io.BytesIO(); im.save(b, "PNG"); b.seek(0); return b


def fonts(run, east="宋体", latin="Times New Roman"):
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None: rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
    for k in ("w:ascii", "w:hAnsi"): rf.set(qn(k), latin)
    rf.set(qn("w:eastAsia"), east)


d = Document()
sec = d.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.top_margin, sec.bottom_margin, sec.left_margin, sec.right_margin = Cm(2.54), Cm(2.54), Cm(3.17), Cm(3.17)
st = d.styles["Normal"]
st.font.size = Pt(10.5); st.font.name = "Times New Roman"
st.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
st.paragraph_format.space_after = Pt(0)
st.paragraph_format.line_spacing = 1.5

# ---- 编号定义 ----
numbering = d.part.numbering_part.element
abs_xml = f'''<w:abstractNum {nsdecls("w")} w:abstractNumId="90"><w:multiLevelType w:val="multilevel"/>
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="chineseCounting"/><w:suff w:val="space"/><w:lvlText w:val="第%1条"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr><w:rPr><w:rFonts w:eastAsia="黑体"/><w:b/></w:rPr></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:isLgl/><w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="425"/></w:tabs><w:ind w:left="425" w:hanging="425"/></w:pPr></w:lvl>
<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%3)"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="851"/></w:tabs><w:ind w:left="851" w:hanging="426"/></w:pPr></w:lvl>
</w:abstractNum>'''
numbering.insert(0, parse_xml(abs_xml))
numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="90"><w:abstractNumId w:val="90"/></w:num>'))
bul = f'''<w:abstractNum {nsdecls("w")} w:abstractNumId="91"><w:multiLevelType w:val="singleLevel"/>
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimalEnclosedCircleChinese"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="420" w:hanging="420"/></w:pPr></w:lvl></w:abstractNum>'''
numbering.insert(1, parse_xml(bul))
numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="91"><w:abstractNumId w:val="91"/></w:num>'))


def numbered(text, lvl, num=90, parts=None):
    p = d.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="{lvl}"/><w:numId w:val="{num}"/></w:numPr>'))
    for part in parts or [text]:
        if isinstance(part, tuple):
            r = p.add_run(part[1]); r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        else:
            r = p.add_run(part)
        fonts(r)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p


# ---- 页眉页脚 ----
hp = sec.header.paragraphs[0]; hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
r = hp.add_run("北京某某律师事务所 · 设备采购合同（草稿）"); fonts(r); r.font.size = Pt(9)
hp._p.get_or_add_pPr().append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>'))
fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
for t in ["第 ", "PAGE", " 页 共 ", "NUMPAGES", " 页"]:
    if t in ("PAGE", "NUMPAGES"):
        fp._p.append(parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr=" {t} "><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>'))
    else:
        r = fp.add_run(t); fonts(r); r.font.size = Pt(9)

# ---- 正文 ----
t = d.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("设备采购合同"); fonts(r, "黑体"); r.bold = True; r.font.size = Pt(22)
for lab in ["甲方（采购方）：", "乙方（供应方）："]:
    p = d.add_paragraph(); r = p.add_run(lab); fonts(r)
    r = p.add_run("北京某某科技有限公司" if "甲" in lab else "上海某某机电设备有限公司"); fonts(r); r.underline = True
p = d.add_paragraph(); p.paragraph_format.first_line_indent = Pt(21)
r = p.add_run("根据《中华人民共和国民法典》及相关法律法规的规定，甲乙双方在平等、自愿、协商一致的基础上，就甲方向乙方采购设备事宜达成如下协议，以资共同遵守。"); fonts(r)
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

S = [
    "乙方应按照本合同附件一所列的规格、型号、数量和质量标准向甲方交付设备，并提供相应的技术资料。",
    "甲方应在收到乙方开具的合法有效发票后三十日内，向乙方支付相应的合同价款，逾期付款的，每逾期一日按应付金额的万分之五支付违约金。",
    "设备的所有权自甲方签收之日起转移至甲方，毁损、灭失的风险自交付之日起由甲方承担，但因乙方原因造成的除外。",
]
heads = ["合同标的", "价款与支付", "交付与验收", "质量保证", "违约责任", "争议解决"]
foot_para = None
for i, h in enumerate(heads):
    numbered(h, 0)
    for j in range(3):
        if i == 1 and j == 0:
            p = numbered("", 1, parts=["合同总价款为人民币", ("ph", "[ ]"), "元（大写：", ("ph", "[ ]"), "），该价款包括设备价款、运输费、安装调试费及税费。"])
            foot_para = p
        else:
            numbered(S[(i + j) % 3], 1)
        if i == 2 and j == 1:
            for k in range(3): numbered(["设备开箱检验记录；", "安装调试报告；", "最终验收证书。"][k], 2)
    if i == 3:
        for k in range(3): numbered(["质保期内免费维修；", "提供备品备件；", "二十四小时内响应。"][k], 0, num=91)

# ---- 修订：插入与删除 ----
p = d.add_paragraph(); r = p.add_run("双方确认，本合同自双方签字盖章之日起生效，"); fonts(r)
p._p.append(parse_xml(f'<w:del {nsdecls("w")} w:id="901" w:author="张律师" w:date="2026-11-01T10:00:00Z"><w:r><w:rPr><w:rFonts w:eastAsia="宋体"/></w:rPr><w:delText>有效期一年。</w:delText></w:r></w:del>'))
p._p.append(parse_xml(f'<w:ins {nsdecls("w")} w:id="902" w:author="张律师" w:date="2026-11-01T10:00:00Z"><w:r><w:rPr><w:rFonts w:eastAsia="宋体"/></w:rPr><w:t>有效期至双方义务履行完毕之日止。</w:t></w:r></w:ins>'))

# ---- 批注 ----
p = d.add_paragraph(); r1 = p.add_run("本合同一式肆份，甲乙双方各执贰份，具有同等法律效力。"); fonts(r1)
d.add_comment(r1, text="请客户确认份数。", author="李律师", initials="LLS")

# ---- 脚注（手工建 footnotes 部件）----
fn_xml = f'''<w:footnotes {nsdecls("w")}>
<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> 含增值税，税率按开票时适用的法定税率计算。</w:t></w:r></w:p></w:footnote>
</w:footnotes>'''.encode("utf-8")
part = Part(PackURI("/word/footnotes.xml"), "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml", fn_xml, d.part.package)
d.part.relate_to(part, RT.FOOTNOTES)
foot_para._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>'))

# ---- 表格：表头底纹、横向 / 纵向合并 ----
d.add_paragraph().add_run("附件一：设备清单").bold = True
tbl = d.add_table(rows=5, cols=4); tbl.style = "Table Grid"
hdr = ["序号", "设备名称", "规格型号", "数量"]
for k, h in enumerate(hdr):
    c = tbl.rows[0].cells[k]; c.text = h
    c._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="D9D9D9"/>'))
rows = [["1", "数控车床", "CK6150", "2"], ["2", "加工中心", "VMC850", "1"], ["3", "三坐标测量机", "CMM-686", "1"], ["合计", "", "", "4"]]
for i, row in enumerate(rows):
    for k, v in enumerate(row): tbl.rows[i + 1].cells[k].text = v
tbl.cell(4, 0).merge(tbl.cell(4, 2))  # 合计行横向合并
tbl.cell(2, 2).merge(tbl.cell(3, 2))  # 纵向合并

# ---- 行内图片 ----
p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run().add_picture(png(240, 120, lambda g: (g.rectangle([2, 2, 237, 117], outline="black", width=3), g.line([10, 100, 120, 30, 230, 80], fill="blue", width=4))), width=Cm(6))

# ---- 浮动图片（公章，压在落款上）----
p = d.add_paragraph(); r = p.add_run("甲方（盖章）："); fonts(r)
seal = png(200, 200, lambda g: (g.ellipse([5, 5, 195, 195], outline=(220, 0, 0, 255), width=8), g.polygon([(100, 60), (112, 95), (150, 95), (120, 117), (131, 152), (100, 130), (69, 152), (80, 117), (50, 95), (88, 95)], fill=(220, 0, 0, 255))))
run = p.add_run(); run.add_picture(seal, width=Cm(4))
inline = run._r.find(".//" + qn("wp:inline"))
anchor = parse_xml(f'''<wp:anchor {nsdecls("wp", "a", "pic", "r")} distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>1080000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>-540000</wp:posOffset></wp:positionV>
<wp:extent cx="1440000" cy="1440000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="77" name="公章"/><wp:cNvGraphicFramePr/></wp:anchor>''')
anchor.append(copy.deepcopy(inline.find(qn("a:graphic"))))
inline.getparent().replace(inline, anchor)
p = d.add_paragraph(); r = p.add_run("日期：2026 年 11 月 22 日"); fonts(r)

# ---- 第二节：横向页面 ----
ns = d.add_section(WD_SECTION.NEW_PAGE)
ns.orientation = WD_ORIENT.LANDSCAPE
ns.page_width, ns.page_height = Cm(29.7), Cm(21)
p = d.add_paragraph(); r = p.add_run("附件二：交货进度表（横向页面）"); fonts(r, "黑体"); r.bold = True
t2 = d.add_table(rows=3, cols=6); t2.style = "Table Grid"
for i in range(3):
    for k in range(6): t2.rows[i].cells[k].text = (["批次", "设备", "数量", "交货日期", "地点", "备注"] if i == 0 else [str(i), "数控车床", "1", f"2027-0{i}-15", "北京", "—"])[k]

d.save(OUT)
print("ok")
