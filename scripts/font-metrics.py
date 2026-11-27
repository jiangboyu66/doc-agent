"""往 src/services/pdfConvert/fontMetrics.json 里写入 Word 字体的字宽表。

用与 Office 字体字宽相同（metric-compatible）的开源字体生成：Cambria ← Caladea，Calibri ← Carlito
（Debian / Ubuntu 的 fonts-crosextra-caladea、fonts-crosextra-carlito）。LibreOffice 打开 docx 时也用它们替换这两种字体。
对照过 Word 导出的 PDF 里嵌入的 Cambria / Calibri 子集：/Widths 与这两种字体逐字相同。

字宽表的格式：{"字体|样式": [[起始码位, [字宽, …]], …]}，样式是 ""、"b"、"i"、"bi"，字宽以 1/1000 em 计，
字体里没有的字符记 -1。码位分段与已有的 Arial 一致。

用法：python3 scripts/font-metrics.py   （需要 fontTools：pip install fonttools）
"""
import json
import os
from fontTools.ttLib import TTFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "src/services/pdfConvert/fontMetrics.json")
DIR = "/usr/share/fonts/truetype/crosextra"
FAMILIES = {
    "Cambria": "Caladea",
    "Calibri": "Carlito",
}
STYLES = {"": "Regular", "b": "Bold", "i": "Italic", "bi": "BoldItalic"}

metrics = json.load(open(OUT, encoding="utf-8"))
segments = [(start, len(ws)) for start, ws in metrics["Arial|"]]
for word, free in FAMILIES.items():
    for key, style in STYLES.items():
        font = TTFont(os.path.join(DIR, f"{free}-{style}.ttf"))
        cmap = font.getBestCmap()
        hmtx = font["hmtx"]
        upm = font["head"].unitsPerEm
        metrics[f"{word}|{key}"] = [
            [start, [round(hmtx[cmap[cp]][0] * 1000 / upm) if cp in cmap else -1 for cp in range(start, start + n)]]
            for start, n in segments
        ]
with open(OUT, "w", encoding="utf-8") as f:
    json.dump(metrics, f, separators=(",", ":"))
print("updated", ", ".join(f"{w}|{k}" for w in FAMILIES for k in STYLES))
