from pathlib import Path
import json,re
from docx import Document
from docx.shared import Pt,Cm,RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT,WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

HERE=Path(__file__).resolve().parent
ROOT=HERE.parents[1]
OUT=ROOT/'outputs/interactive-teacher-20261025'
OUT.mkdir(parents=True,exist_ok=True)
md=(HERE/'report-body.txt').read_text().strip()
doc=Document();sec=doc.sections[0]
sec.page_width=Cm(21.59);sec.page_height=Cm(27.94)
sec.top_margin=Cm(1.8);sec.bottom_margin=Cm(1.8);sec.left_margin=Cm(1.9);sec.right_margin=Cm(1.9)
sec.header_distance=Cm(.7);sec.footer_distance=Cm(.7)
for name in ['Normal','Title','Heading 1','Heading 2','Heading 3']:
 s=doc.styles[name];s.font.name='Noto Sans CJK SC';s.font.color.rgb=RGBColor(0,0,0)
 f=s._element.get_or_add_rPr().get_or_add_rFonts()
 for k in list(f.attrib):
  if 'theme' in k.lower():del f.attrib[k]
 for tag in ['ascii','hAnsi','eastAsia']:f.set(qn('w:'+tag),'Noto Sans CJK SC')
 for b in s._element.xpath('./w:pPr/w:pBdr'):b.getparent().remove(b)
s=doc.styles['Normal'];s.font.size=Pt(11);s.paragraph_format.line_spacing=Pt(16.5);s.paragraph_format.space_after=Pt(7);s.paragraph_format.widow_control=True
doc.styles['Title'].font.size=Pt(22);doc.styles['Title'].paragraph_format.space_after=Pt(13)
for n,size in [(1,15),(2,12),(3,11)]:
 s=doc.styles[f'Heading {n}'];s.font.size=Pt(size);s.font.bold=True;s.paragraph_format.space_before=Pt(11);s.paragraph_format.space_after=Pt(6);s.paragraph_format.keep_with_next=True
p=sec.header.paragraphs[0];p.text='金种子杯  互动数字教师与长理产业路线比较';p.runs[0].font.size=Pt(8)
p=sec.footer.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.RIGHT;p.add_run('2026年10月25日  ·  ')
fld=OxmlElement('w:fldSimple');fld.set(qn('w:instr'),'PAGE');p._p.append(fld)
for r in p.runs:r.font.size=Pt(8)
def inline(p,text,size=None):
 for item in re.split(r'(\*\*.*?\*\*|\[[^\]]+\]\(https?://[^)]+\))',text):
  if not item:continue
  if item.startswith('[') and '](' in item:
   label,url=re.match(r'\[([^\]]+)\]\(([^)]+)\)',item).groups();h=OxmlElement('w:hyperlink');h.set(qn('r:id'),p.part.relate_to(url,RT.HYPERLINK,is_external=True));r=OxmlElement('w:r');pr=OxmlElement('w:rPr');color=OxmlElement('w:color');color.set(qn('w:val'),'245A81');pr.append(color)
   if size:
    sz=OxmlElement('w:sz');sz.set(qn('w:val'),str(int(size*2)));pr.append(sz)
   r.append(pr);t=OxmlElement('w:t');t.text=label;r.append(t);h.append(r);p._p.append(h)
  else:
   r=p.add_run(item[2:-2] if item.startswith('**') else item);r.bold=item.startswith('**')
   if size:r.font.size=Pt(size)
def table(lines):
 grid=[[v.strip() for v in l.strip('|').split('|')] for l in lines if not re.match(r'^\|[-: |]+\|$',l)]
 n=len(grid[0]);isappend=grid[0][0]=='项目与年份'
 widths=[3.5,5.3,6.7,1.7] if isappend else {2:[5,12.4],3:[3.6,6.9,6.9],4:[2.3,5.0,5.0,5.1]}[n]
 if grid[0][0]=='方案表述':widths=[5.7,2.4,9.3]
 t=doc.add_table(rows=0,cols=n);t.autofit=False;t.alignment=WD_TABLE_ALIGNMENT.CENTER
 for col,w in zip(t.columns,widths):col.width=Cm(w)
 for ri,vals in enumerate(grid):
  cells=t.add_row().cells;tr=t.rows[-1]._tr.get_or_add_trPr();tr.append(OxmlElement('w:cantSplit'))
  if ri==0:tr.append(OxmlElement('w:tblHeader'))
  for ci,(cell,val,w) in enumerate(zip(cells,vals,widths)):
   cell.width=Cm(w);cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
   pr=cell._tc.get_or_add_tcPr();border=OxmlElement('w:tcBorders')
   for edge in ['top','left','bottom','right']:
    e=OxmlElement('w:'+edge);e.set(qn('w:val'),'single');e.set(qn('w:sz'),'4');e.set(qn('w:color'),'D9D9D9');border.append(e)
   pr.append(border);shade=OxmlElement('w:shd');shade.set(qn('w:fill'),'DCE6EE' if ri==0 else ('F5F7F9' if ri%2==0 else 'FFFFFF'));pr.append(shade)
   mar=OxmlElement('w:tcMar')
   for edge in ['top','left','bottom','right']:
    e=OxmlElement('w:'+edge);e.set(qn('w:w'),'85');e.set(qn('w:type'),'dxa');mar.append(e)
   pr.append(mar);p=cell.paragraphs[0];p.paragraph_format.space_after=Pt(1);p.paragraph_format.line_spacing=Pt(14);inline(p,val,9.5)
   if grid[0][0]=='方案表述' and ci==1:p.alignment=WD_ALIGN_PARAGRAPH.CENTER
   for r in p.runs:r.bold=ri==0
 doc.add_paragraph().paragraph_format.space_after=Pt(0)
lines=md.splitlines();i=0
while i<len(lines):
 line=lines[i]
 if not line.strip():i+=1;continue
 if line.startswith('|'):
  group=[]
  while i<len(lines) and lines[i].startswith('|'):group.append(lines[i]);i+=1
  table(group);continue
 if line.startswith('# '):doc.add_paragraph('互动数字教师\n与长理产业路线比较',style='Title')
 elif line.startswith('## '):
  p=doc.add_heading(line[3:],1)
  p.paragraph_format.page_break_before=False
 elif line.startswith('### '):doc.add_heading(line[4:],2)
 else:inline(doc.add_paragraph(),line)
 i+=1
doc.core_properties.title='互动数字教师与长理产业路线比较'
doc.core_properties.subject='需求 创新 实现程度与参赛竞争力的比较研究'
doc.core_properties.author=''
path=OUT/'互动数字教师与长理产业路线比较.docx';doc.save(path)
assert '' not in md
print(json.dumps({'output':str(path),'characters':len(md),'distinct_links':len(set(re.findall(r'\]\((https?://[^)]+)\)',md)))},ensure_ascii=False))
