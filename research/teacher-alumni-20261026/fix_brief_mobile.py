from pathlib import Path
from copy import deepcopy
import hashlib
import json
import re

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.text.paragraph import Paragraph
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
OUT = ROOT / 'outputs/teacher-alumni-20261026'
original = OUT / '03_第三届金种子杯选题汇报_精简版.docx'
target = OUT / '03_第三届金种子杯选题汇报_精简版_手机兼容修订.docx'
doc = Document(original)

def text_content(document):
    return re.sub(r'\s+', '', ''.join(document.element.body.xpath('.//w:t/text()')))

before = text_content(doc)
table_headers = ''.join(c.text for c in doc.tables[0].rows[0].cells)
FONT = 'Noto Sans CJK SC'

def font_properties(rpr, size=None, bold=None):
    fonts = rpr.find(qn('w:rFonts'))
    if fonts is None:
        fonts = OxmlElement('w:rFonts')
        rpr.insert(0, fonts)
    fonts.attrib.clear()
    for script in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
        fonts.set(qn('w:' + script), FONT)
    if size is not None:
        for tag in ['sz', 'szCs']:
            e = rpr.find(qn('w:' + tag))
            if e is None:
                e = OxmlElement('w:' + tag)
                rpr.append(e)
            e.set(qn('w:val'), str(round(size * 2)))
    lang = rpr.find(qn('w:lang'))
    if lang is None:
        lang = OxmlElement('w:lang')
        rpr.append(lang)
    lang.set(qn('w:val'), 'zh-CN')
    lang.set(qn('w:eastAsia'), 'zh-CN')

# Use explicit minimum line heights, avoiding fixed line boxes and
# reliance on inherited theme font metrics in the viewer.
for name, size in [('Normal', 11), ('Title', 22), ('Subtitle', 11),
                   ('Heading 1', 16), ('Heading 2', 13), ('Footer', 10)]:
    style = doc.styles[name]
    font_properties(style._element.get_or_add_rPr(), size)
    style.font.color.rgb = RGBColor(0, 0, 0)
    minimum = {'Normal': 17, 'Title': 28, 'Subtitle': 16,
               'Heading 1': 20, 'Heading 2': 18, 'Footer': 13}[name]
    style.paragraph_format.line_spacing = Pt(minimum)
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
    snap = OxmlElement('w:snapToGrid')
    snap.set(qn('w:val'), '0')
    style._element.get_or_add_pPr().append(snap)
    if name == 'Subtitle':
        style.font.italic = False
        style.paragraph_format.space_after = Pt(10)
    elif name == 'Normal':
        style.paragraph_format.space_after = Pt(5)
    elif name.startswith('Heading'):
        style.paragraph_format.space_before = Pt(8)
        style.paragraph_format.space_after = Pt(6)

defaults = doc.styles.element.find(qn('w:docDefaults'))
font_properties(defaults.find(qn('w:rPrDefault')).find(qn('w:rPr')), 11)
spacing = defaults.find(qn('w:pPrDefault')).find(qn('w:pPr')).find(qn('w:spacing'))
spacing.set(qn('w:line'), '340')
spacing.set(qn('w:lineRule'), 'atLeast')
spacing.set(qn('w:after'), '140')

for p in doc.paragraphs:
    # A valid paragraph-end run size also matters to cell/paragraph height
    # in lightweight renderers. The desktop style still controls appearance.
    ppr = p._p.get_or_add_pPr()
    mark = ppr.find(qn('w:rPr'))
    if mark is None:
        mark = OxmlElement('w:rPr')
        ppr.append(mark)
    size = p.style.font.size.pt if p.style.font.size else 11
    font_properties(mark, size)
    for r in p.runs:
        font_properties(r._r.get_or_add_rPr(), r.font.size.pt if r.font.size else size)

for table in doc.tables:
    table.autofit = False
    total = sum(col.width.twips for col in table.columns)
    width = table._tbl.tblPr.find(qn('w:tblW'))
    width.set(qn('w:type'), 'dxa')
    width.set(qn('w:w'), str(total))
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            # Top alignment avoids viewer-specific center/baseline arithmetic.
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            content = cell.text
            cell.text = content.replace('\n', '  ')
            p = cell.paragraphs[0]
            pf = p.paragraph_format
            pf.line_spacing = Pt(16)
            pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
            pf.space_before = Pt(0)
            pf.space_after = Pt(2)
            pf.keep_with_next = ri == 0
            pf.keep_together = True
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            mark = OxmlElement('w:rPr')
            font_properties(mark, 10.5)
            p._p.get_or_add_pPr().append(mark)
            for r in p.runs:
                font_properties(r._r.get_or_add_rPr(), 10.5)
                r.font.size = Pt(10.5)
                r.bold = ri == 0
            props = cell._tc.get_or_add_tcPr()
            margins = props.find(qn('w:tcMar'))
            for child in list(margins):
                child.set(qn('w:w'), '100')
            # Keep the schema order canonical instead of relying on a
            # tolerant desktop parser to accept out-of-order properties.
            order = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge',
                     'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection',
                     'tcFitText', 'vAlign', 'hideMark', 'headers', 'cellIns',
                     'cellDel', 'cellMerge', 'tcPrChange']
            rank = {qn('w:' + tag): i for i, tag in enumerate(order)}
            props[:] = sorted(props, key=lambda e: rank.get(e.tag, 999))
            for borders_tag in ['tcBorders', 'tcMar']:
                edges = props.find(qn('w:' + borders_tag))
                edge_order = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl']
                erank = {qn('w:' + tag): i for i, tag in enumerate(edge_order)}
                edges[:] = sorted(edges, key=lambda e: erank.get(e.tag, 999))
    # Remove the 3 pt empty spacer paragraph that can collapse in the mobile
    # viewer; use the following heading's ordinary paragraph spacing.
    following = table._tbl.getnext()
    if following is not None and following.tag == qn('w:p') and not ''.join(following.itertext()).strip():
        following.getparent().remove(following)

    # Present the same five comparisons in the normal document flow. This
    # removes the failing mobile table layout path rather than relying on
    # Tencent's renderer to interpret cell baselines and heights correctly.
    for row in table.rows[1:]:
        values = [c.text for c in row.cells]
        e = OxmlElement('w:p')
        table._tbl.addprevious(e)
        p = Paragraph(e, doc._body)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.keep_together = True
        for ci, value in enumerate(values):
            if ci:
                p.add_run('：' if ci == 1 else '  ')
            if ci == 2:
                value = value.replace('  ', '；', 1)
            r = p.add_run(value)
            r.bold = ci == 0
            font_properties(r._r.get_or_add_rPr(), 11)
    table._tbl.getparent().remove(table._tbl)

# Replace simple fields with standard complex field markup and cached text.
# This keeps page numbers editable and avoids the screenshot's bare slash.
for section in doc.sections:
    p = section.footer.paragraphs[0]
    p.clear()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    for instruction, cached in [('PAGE', '1'), ('NUMPAGES', '3')]:
        if instruction == 'NUMPAGES':
            p.add_run(' / ')
        start = p.add_run()
        e = OxmlElement('w:fldChar')
        e.set(qn('w:fldCharType'), 'begin')
        start._r.append(e)
        r = p.add_run()
        e = OxmlElement('w:instrText')
        e.set(qn('xml:space'), 'preserve')
        e.text = ' ' + instruction + ' '
        r._r.append(e)
        r = p.add_run()
        e = OxmlElement('w:fldChar')
        e.set(qn('w:fldCharType'), 'separate')
        r._r.append(e)
        p.add_run(cached)
        r = p.add_run()
        e = OxmlElement('w:fldChar')
        e.set(qn('w:fldCharType'), 'end')
        r._r.append(e)
    for r in p.runs:
        font_properties(r._r.get_or_add_rPr(), 10)

original_without_headers = before.replace(re.sub(r'\s+', '', table_headers), '', 1)
assert re.sub('[：；]', '', original_without_headers) == re.sub('[：；]', '', text_content(doc)), 'Proposal wording changed'
assert not doc.element.body.xpath('.//w:spacing[@w:lineRule="exact"]')
assert not doc.element.body.xpath('.//w:tbl//w:br')
assert not doc.element.body.xpath('.//w:tbl//w:vAlign[@w:val="center"]')
doc.save(target)
report = {
    'source': str(original), 'output': str(target),
    'all_proposal_wording_preserved': True,
    'table_column_labels_removed': ['方向', '我们想做什么', '当前位置与关键条件'],
    'source_sha256': hashlib.sha256(original.read_bytes()).hexdigest(),
    'output_sha256': hashlib.sha256(target.read_bytes()).hexdigest(),
    'changes': ['minimum line spacing permitting automatic expansion',
                'comparison table replaced with five normal-flow paragraphs',
                'all proposal descriptions, priorities and resource conditions preserved',
                'explicit font properties',
                'normal paragraph gaps',
                'complex page number fields with cached values'],
    'wechat_device_test': 'Not available in this environment',
}
(HERE / 'brief-mobile-fix-check.json').write_text(json.dumps(report, ensure_ascii=False, indent=2))
print(json.dumps(report, ensure_ascii=False, indent=2))
