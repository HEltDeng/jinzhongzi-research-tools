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
OUT=ROOT/'outputs/jzz-reality-ai-20261025'
OUT.mkdir(parents=True,exist_ok=True)
previous=json.loads((ROOT/'research/gold-study-20261025/case_analysis.json').read_text())
mat=[l.split('|') for l in (HERE/'maturity.txt').read_text().splitlines() if l]
assert len(mat)==len(previous['cases'])==36
md=(HERE/'report-body.txt').read_text().strip()
md+='\n\n## 附录 两届36项金奖的成熟度证据备注\n\n'
md+='本表按上一轮正式名单顺序逐项复核，描述公开资料支持到的程度，不是技术成熟度认证或真实性审计。来源链接直接对应项目材料；缺少资料的项目标为未知。上一轮A、B、C是证据档位，也不能直接转换为产品阶段。\n\n'
table=['|项目与年份|可见进展|时间与适用边界|来源|','|---|---|---|---|']
for c,(state,limit) in zip(previous['cases'],mat):
 refs=c['source_ids']
 urls=[previous['sources'][r]['url'] for r in refs]
 if c['id']==16:urls=['https://m.voc.com.cn/xhn/news/202505/29503302.html']
 if c['id']==3:urls.append('https://www.hhjt.com/info/1011/5871.htm')
 links=' '.join(f'[资料{j}]({u})' for j,u in enumerate(dict.fromkeys(urls),1))
 table.append(f"|{c['id']:02d} {c['short_name']} {c['year']}|{state}|{limit}|{links}|")
md+='\n'.join(table)+'\n\n名单依据：[2025年正式名单](https://hunan.gov.cn/topic/dxscy/dxscyZCJC/dxscysjzc/202510/t20251015_33826774.html)、[2026年正式名单](https://jyt.hunan.gov.cn/jyt/sjyt/xxgk/tzgg/202607/t20260708_1180348.html)。具体全名、学校及赛道沿用前篇逐项研究。\n'
(HERE/'render-input.txt').write_text(md)
doc=Document();sec=doc.sections[0]
sec.page_width=Cm(21);sec.page_height=Cm(29.7)
sec.top_margin=Cm(1.8);sec.bottom_margin=Cm(1.8);sec.left_margin=Cm(1.9);sec.right_margin=Cm(1.9)
sec.header_distance=Cm(.7);sec.footer_distance=Cm(.7)
for name in ['Normal','Title','Heading 1','Heading 2','Heading 3']:
 s=doc.styles[name];s.font.name='Noto Sans CJK SC';s.font.color.rgb=RGBColor(0,0,0)
 f=s._element.get_or_add_rPr().get_or_add_rFonts()
 for k in list(f.attrib):
  if 'theme' in k.lower():del f.attrib[k]
 for tag in ['ascii','hAnsi','eastAsia']:f.set(qn('w:'+tag),'Noto Sans CJK SC')
 for b in s._element.xpath('./w:pPr/w:pBdr'):b.getparent().remove(b)
s=doc.styles['Normal'];s.font.size=Pt(10.5);s.paragraph_format.line_spacing=Pt(15.5);s.paragraph_format.space_after=Pt(7);s.paragraph_format.widow_control=True
doc.styles['Title'].font.size=Pt(22);doc.styles['Title'].paragraph_format.space_after=Pt(13)
for n,size in [(1,15),(2,12),(3,11)]:
 s=doc.styles[f'Heading {n}'];s.font.size=Pt(size);s.font.bold=True;s.paragraph_format.space_before=Pt(11);s.paragraph_format.space_after=Pt(6);s.paragraph_format.keep_with_next=True
p=sec.header.paragraphs[0];p.text='金种子杯  获奖项目成熟度与AI选题判断';p.runs[0].font.size=Pt(8)
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
 widths=[3.5,5.3,6.7,1.7] if isappend else {2:[5,12.2],3:[3.1,7.4,6.7],4:[3.6,4.5,5.7,3.4]}[n]
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
   pr.append(mar);p=cell.paragraphs[0];p.paragraph_format.space_after=Pt(1);p.paragraph_format.line_spacing=Pt(13);inline(p,val,9)
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
 if line.startswith('# '):doc.add_paragraph('金种子杯获奖项目成熟度\n与AI选题判断',style='Title')
 elif line.startswith('## '):
  p=doc.add_heading(line[3:],1)
  if line.startswith(('## 二','## 三','## 四','## 五','## 六','## 七','## 八','## 附录')):p.paragraph_format.page_break_before=True
 elif line.startswith('### '):doc.add_heading(line[4:],2)
 else:inline(doc.add_paragraph(),line)
 i+=1
doc.core_properties.title='金种子杯获奖项目成熟度与AI选题判断'
doc.core_properties.subject='周一师生讨论依据 36项成熟度备注与三类候选方向'
doc.core_properties.author=''
path=OUT/'金种子杯获奖项目成熟度与AI选题判断.docx';doc.save(path)
assert '' not in md and len(mat)==36
print(json.dumps({'output':str(path),'characters':len(md),'cases':len(mat),'distinct_links':len(set(re.findall(r'\]\((https?://[^)]+)\)',md)))},ensure_ascii=False))
