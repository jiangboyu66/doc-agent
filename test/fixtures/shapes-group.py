"""生成测试夹具 shapes-group.doc：一个形状组合（绿色矩形、带箭头的连接线、自由曲线三角形）+ 页脚图文框页码。
本脚本生成 .docx，再用 LibreOffice 另存为 .doc：
  python shapes-group.py /tmp/shapes.docx
  soffice --headless --convert-to doc:"MS Word 97" --outdir /tmp /tmp/shapes.docx   # → /tmp/shapes.doc
LibreOffice 写 .doc 时：自由曲线存成"无预设几何 + 顶点"，几何数组的长度不含 6 字节数组头（解析器要兼容）。
"""
import zipfile, sys
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
def wsp(i, x, y, cx, cy, geom, text=None, line=False):
    tx = f'<wps:txbx><w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>' if text else ''
    nv = '<wps:cNvCnPr/>' if line else '<wps:cNvSpPr/>'
    ln = '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>' + ('<a:tailEnd type="triangle"/>' if line else '') + '</a:ln>'
    fill = '<a:noFill/>' if line else '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>'
    return f'<wps:wsp><wps:cNvPr id="{i}" name="s{i}"/>{nv}<wps:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>{geom}{fill}{ln}</wps:spPr>{tx}<wps:bodyPr/></wps:wsp>'
rect = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
conn = '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
free = '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:lnTo><a:pt x="50" y="0"/></a:lnTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>'
cx, cy = 4572000, 2286000
group = (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
  f'<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
  f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/><wp:docPr id="1" name="g"/><wp:cNvGraphicFramePr/>'
  f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/><a:chOff x="0" y="0"/><a:chExt cx="{cx}" cy="{cy}"/></a:xfrm></wpg:grpSpPr>'
  + wsp(2, 0, 0, 1828800, 914400, rect, "流程开始") + wsp(3, 1828800, 457200, 914400, 0, conn, line=True) + wsp(4, 2743200, 0, 1828800, 1828800, free)
  + '</wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')
body = f'<w:p><w:r><w:t>示意图如下。</w:t></w:r></w:p><w:p>{group}</w:p><w:p><w:r><w:t>图后的正文。</w:t></w:r></w:p>'
sect = '<w:sectPr><w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'
ftr = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {W}><w:p><w:pPr><w:framePr w:w="1456" w:wrap="around" w:vAnchor="text" w:hAnchor="page" w:x="9022" w:y="97"/><w:jc w:val="right"/></w:pPr>'
  '<w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
  '<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p></w:ftr>')
ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
drels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>'
with zipfile.ZipFile(sys.argv[1], 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', rels); z.writestr('word/document.xml', doc)
    z.writestr('word/_rels/document.xml.rels', drels); z.writestr('word/footer1.xml', ftr)
