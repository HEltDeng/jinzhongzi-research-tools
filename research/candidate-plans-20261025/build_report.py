from pathlib import Path
import json, re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
OUT = ROOT / 'outputs/candidate-plans-20261025'
OUT.mkdir(parents=True, exist_ok=True)
content = (HERE / 'report-body.txt').read_text().strip()
doc = Document()
sec = doc.sections[0]
sec.page_width = Cm(21.59)
sec.page_height = Cm(27.94)
sec.top_margin = Cm(1.8)
sec.bottom_margin = Cm(1.8)
sec.left_margin = Cm(1.9)
sec.right_margin = Cm(1.9)
sec.footer_distance = Cm(.75)
font = 'Noto Sans CJK SC'
for name in ['Normal', 'Title', 'Heading 1', 'Heading 2', 'Heading 3']:
    s = doc.styles[name]
    s.font.name = font
    s.font.color.rgb = RGBColor(0, 0, 0)
    f = s._element.get_or_add_rPr().get_or_add_rFonts()
    for k in list(f.attrib):
        if 'theme' in k.lower():
            del f.attrib[k]
    for tag in ['ascii', 'hAnsi', 'eastAsia']:
        f.set(qn('w:' + tag), font)
    for b in s._element.xpath('./w:pPr/w:pBdr'):
        b.getparent().remove(b)
s = doc.styles['Normal']
s.font.size = Pt(11)
s.paragraph_format.line_spacing = Pt(16.5)
s.paragraph_format.space_after = Pt(7)
s.paragraph_format.widow_control = True
doc.styles['Title'].font.size = Pt(22)
doc.styles['Title'].paragraph_format.space_after = Pt(14)
for n, size in [(1, 15), (2, 12), (3, 11)]:
    s = doc.styles[f'Heading {n}']
    s.font.size = Pt(size)
    s.font.bold = True
    s.paragraph_format.space_before = Pt(11)
    s.paragraph_format.space_after = Pt(6)
    s.paragraph_format.keep_with_next = True
p = sec.footer.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
p.paragraph_format.space_after = Pt(0)
fld = OxmlElement('w:fldSimple')
fld.set(qn('w:instr'), 'PAGE')
r = OxmlElement('w:r')
pr = OxmlElement('w:rPr')
sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '20'); pr.append(sz)
col = OxmlElement('w:color'); col.set(qn('w:val'), '000000'); pr.append(col)
r.append(pr)
t = OxmlElement('w:t'); t.text = '1'; r.append(t)
fld.append(r); p._p.append(fld)

def inline(p, text, size=None):
    for item in re.split(r'(\*\*.*?\*\*|\[[^\]]+\]\(https?://[^)]+\))', text):
        if not item:
            continue
        if item.startswith('[') and '](' in item:
            label, url = re.match(r'\[([^\]]+)\]\(([^)]+)\)', item).groups()
            h = OxmlElement('w:hyperlink')
            h.set(qn('r:id'), p.part.relate_to(url, RT.HYPERLINK, is_external=True))
            r = OxmlElement('w:r'); pr = OxmlElement('w:rPr')
            color = OxmlElement('w:color'); color.set(qn('w:val'), '000000'); pr.append(color)
            under = OxmlElement('w:u'); under.set(qn('w:val'), 'single'); pr.append(under)
            if size:
                sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(int(size * 2))); pr.append(sz)
            r.append(pr); t = OxmlElement('w:t'); t.text = label
            r.append(t); h.append(r); p._p.append(h)
        else:
            r = p.add_run(item[2:-2] if item.startswith('**') else item)
            r.bold = item.startswith('**')
            r.font.color.rgb = RGBColor(0, 0, 0)
            if size:
                r.font.size = Pt(size)

def table(lines):
    grid = [[v.strip() for v in l.strip('|').split('|')] for l in lines if not re.match(r'^\|[-: |]+\|$', l)]
    n = len(grid[0])
    assert all(len(row) == n for row in grid)
    widths = {3: [3.0, 7.3, 7.3], 4: [3.0, 4.8, 4.8, 5.0], 7: [4.1, 2.25, 2.25, 2.25, 2.25, 2.25, 2.25]}[n]
    if grid[0][0] == '方案' and grid[0][1] == '做什么产品':
        widths = [3.1, 5.5, 4.1, 4.9]
    if grid[0][0] == '方案' and grid[0][1] == '控制额':
        widths = [1.4, 2.1, 7.0, 7.1]
    if grid[0][0] == '评审项':
        widths = [4.3, 1.9, 11.4]
    if grid[0][0] == '时间':
        widths = [3.7, 7.1, 6.8]
    if n == 7:
        widths = [4.0] + [2.27] * 6
    tb = doc.add_table(rows=0, cols=n)
    tb.autofit = False
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c, w in zip(tb.columns, widths):
        c.width = Cm(w)
    for ri, vals in enumerate(grid):
        cells = tb.add_row().cells
        tr = tb.rows[-1]._tr.get_or_add_trPr()
        tr.append(OxmlElement('w:cantSplit'))
        if ri == 0:
            tr.append(OxmlElement('w:tblHeader'))
        for ci, (cell, val, w) in enumerate(zip(cells, vals, widths)):
            cell.width = Cm(w)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            pr = cell._tc.get_or_add_tcPr()
            border = OxmlElement('w:tcBorders')
            for edge in ['top', 'left', 'bottom', 'right']:
                e = OxmlElement('w:' + edge)
                e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '4'); e.set(qn('w:color'), 'D9D9D9')
                border.append(e)
            pr.append(border)
            shade = OxmlElement('w:shd'); shade.set(qn('w:fill'), 'E4ECF3' if ri == 0 else 'FFFFFF'); pr.append(shade)
            mar = OxmlElement('w:tcMar')
            for edge in ['top', 'left', 'bottom', 'right']:
                margin = '50' if n == 7 and edge in ['top', 'bottom'] else '80'
                e = OxmlElement('w:' + edge); e.set(qn('w:w'), margin); e.set(qn('w:type'), 'dxa'); mar.append(e)
            pr.append(mar)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing = Pt(13.5 if n == 7 else 14.5)
            if len(grid) <= 7:
                p.paragraph_format.keep_with_next = ri < len(grid) - 1
            inline(p, val, 10)
            if n == 7 and ci > 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = ri == 0
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(3)
    p.paragraph_format.space_before = Pt(0)

lines = content.splitlines()
i = 0
while i < len(lines):
    line = lines[i]
    if not line.strip():
        i += 1; continue
    if line.startswith('|'):
        group = []
        while i < len(lines) and lines[i].startswith('|'):
            group.append(lines[i]); i += 1
        table(group); continue
    if line.startswith('# '):
        doc.add_paragraph(line[2:], style='Title')
    elif line.startswith('## '):
        p = doc.add_heading(line[3:], 1)
        p.paragraph_format.page_break_before = line.startswith('## 二 ')
    elif line.startswith('### '):
        doc.add_heading(line[4:], 2)
    else:
        p = doc.add_paragraph()
        inline(p, line)
        following = next((s for s in lines[i + 1:] if s.strip()), '')
        if following.startswith('**可借鉴：**'):
            p.paragraph_format.keep_with_next = True
    i += 1

doc.core_properties.title = '金种子杯候选项目方案与推进计划'
doc.core_properties.subject = '项目落地 创新评分 资源评估 往届金奖参照'
doc.core_properties.author = ''
path = OUT / '金种子杯候选项目方案与推进计划.docx'
doc.save(path)
assert '' not in content
print(json.dumps({'output': str(path), 'characters': len(content), 'unique_sources': len(set(re.findall(r'\]\((https?://[^)]+)\)', content)))}, ensure_ascii=False))
