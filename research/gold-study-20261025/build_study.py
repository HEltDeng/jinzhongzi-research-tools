from pathlib import Path
import json, re, collections
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

ROOT=Path(__file__).resolve().parents[2]
HERE=Path(__file__).resolve().parent
OUT=ROOT/'outputs/jzz-gold-study-20261025'
OUT.mkdir(parents=True,exist_ok=True)
gold=json.loads((ROOT/'research/gold_analysis.json').read_text())
rows=[l.split('|') for l in (HERE/'cases.txt').read_text().splitlines() if l]
sources={p[0]:dict(zip(['title','date','url'],p[1:])) for p in (l.split('|') for l in (HERE/'sources.txt').read_text().splitlines() if l)}
assert len(rows)==len(gold)==36
data=[]
for i,(r,g) in enumerate(zip(rows,gold),1):
    assert len(r)==10
    name,category,cross,score,level,refs,fact,reason,creative,gap=r
    scores=list(map(int,score.split(','))); assert len(scores)==4 and all(1<=s<=5 for s in scores)
    ids=refs.split(',');assert all(s in sources for s in ids)
    full=g['项目名称']
    full={18:'SiC High-Temp Ion Implantation Equipment R&D & Industrialization',19:'MetaSample 64G: Direct Sampling SoC Innovation',30:'AI-Biomedical Metamaterials R&D'}.get(i,full)
    data.append({'id':i,'year':g['获奖年份'],'official_track':g['官方赛道'],'official_name':g['项目名称'],'display_name':full,'short_name':name,'schools':g['所属学校'],'category':category,'combination':cross,'scores':dict(zip(['I','C','D','V'],scores)),'total':sum(scores)*5,'evidence_level':level,'source_ids':ids,'facts':fact,'award_analysis_inference':reason,'creativity_analysis':creative,'gaps':gap,'official_list_page':g['名单页码'],'official_local_attachment':g['原始附件']})
assert collections.Counter(d['year'] for d in data)=={2025:16,2026:20}
assert collections.Counter(d['evidence_level'] for d in data)=={'A':10,'B':13,'C':13}

def reflinks(ids):
    return '、'.join(f"[{s} {sources[s]['title']}]({sources[s]['url']})" for s in dict.fromkeys(ids))

parts=[(HERE/'intro.md').read_text().strip()]
year=None;track=None
for d in data:
    if d['year']!=year:
        year=d['year'];track=None
        parts.append('## 五 2025年16个金奖项目逐项分析' if year==2025 else '## 六 2026年20个金奖项目逐项分析')
    if d['official_track']!=track:
        track=d['official_track']; parts.append('### '+track.replace('+','加'))
        if '专项' in track:parts.append('以下两项为AI 3D大模型应用专项赛金奖。本研究采用相同创意量表便于拆解方案，但不将专项赛与主赛道视为同一竞争池，也未取得专项赛独立评分细表。')
    parts.append(f"#### {d['id']:02d} {d['short_name']}")
    parts.append(f"**官方项目名**：{d['display_name']}。**学校**：{d['schools']}。**官方赛道**：{track}。")
    label={'A':'研究评分','B':'条件评分','C':'概念暂评'}[d['evidence_level']]
    sc=d['scores']; scoretext=' / '.join(str(sc[k]) for k in ['I','C','D','V'])
    parts.append(f"**结合方式**：{d['combination']}。**创意{label}**：{d['total']}/100，分项I/C/D/V为{scoretext}，每项满分5。**证据档位**：{d['evidence_level']}。")
    parts.append('**已核信息**：'+d['facts'])
    parts.append('**获奖原因分析**：'+d['award_analysis_inference'])
    parts.append('**创意表现与评分依据**：'+d['creativity_analysis'])
    parts.append('**尚待核实**：'+d['gaps'])
    parts.append('依据：'+reflinks(['R01' if year==2025 else 'R02']+d['source_ids'])+'。')
parts.append('### 36项创意评分对照')
parts.append('I为问题洞察，C为结合逻辑，D为方案差异，V为价值路径。各项满分5，总分满分100。按正式名单顺序排列，不作获奖名次排序。A为研究评分，B为条件评分，C为概念暂评；C档尤其不能与A档据分数判断真实优劣。')
table=['|序号与简称|I|C|D|V|总分|证据|','|---|---|---|---|---|---|---|']
for d in data:table.append('|'+ '|'.join([f"{d['id']:02d} {d['short_name']}"]+[str(d['scores'][k]) for k in ['I','C','D','V']]+[str(d['total']),d['evidence_level']])+'|')
parts.append('\n'.join(table))
parts.append((HERE/'closing.md').read_text().strip())
parts.append('## 十 来源索引')
parts.append('以下为正文实际采用的35项来源。名单来源确认奖项与赛道；案例来源支撑项目介绍，可信范围和时间限制已写入对应分析。所有网页链接均来自本轮检索或工作区已有公开来源归档，不含私人资料。')
for rid,s in sources.items():parts.append(f"[{rid} {s['title']}]({s['url']})。时间及限制：{s['date']}。")
md='\n\n'.join(parts)+'\n'
mdpath=OUT/'2025与2026年金种子杯36项金奖深度研究.md'
mdpath.write_text(md)
(HERE/'case_analysis.json').write_text(json.dumps({'as_of':'2026-10-25','method':'I/C/D/V each 1–5; total = 5 × sum; independent research not official scores','cases':data,'sources':sources},ensure_ascii=False,indent=2))

doc=Document(); sec=doc.sections[0]
sec.page_width=Cm(21);sec.page_height=Cm(29.7)
sec.top_margin=Cm(1.7);sec.bottom_margin=Cm(1.7);sec.left_margin=Cm(1.9);sec.right_margin=Cm(1.9)
sec.header_distance=Cm(.7);sec.footer_distance=Cm(.7)
for name in ['Normal','Title','Subtitle','Heading 1','Heading 2','Heading 3','Heading 4']:
    sty=doc.styles[name];sty.font.name='Noto Sans CJK SC';sty.font.color.rgb=RGBColor(0,0,0)
    rf=sty._element.get_or_add_rPr().get_or_add_rFonts()
    for k in list(rf.attrib):
        if 'theme' in k.lower():del rf.attrib[k]
    for tag in ['ascii','hAnsi','eastAsia']:rf.set(qn('w:'+tag),'Noto Sans CJK SC')
    for b in sty._element.xpath('./w:pPr/w:pBdr'):b.getparent().remove(b)
norm=doc.styles['Normal'];norm.font.size=Pt(10.5)
norm.paragraph_format.line_spacing=Pt(15.5);norm.paragraph_format.space_after=Pt(6)
norm.paragraph_format.widow_control=True
doc.styles['Title'].font.size=Pt(23)
doc.styles['Title'].paragraph_format.space_after=Pt(14)
for n,size in [(1,16),(2,13),(3,11.5)]:
    sty=doc.styles[f'Heading {n}'];sty.font.size=Pt(size);sty.font.bold=True
    sty.paragraph_format.space_before=Pt(12 if n<3 else 9);sty.paragraph_format.space_after=Pt(6)
    sty.paragraph_format.keep_with_next=True
header=sec.header.paragraphs[0];header.text='湖南金种子杯  2025与2026年金奖项目研究'
header.runs[0].font.size=Pt(8);header.runs[0].font.color.rgb=RGBColor(0,0,0)
footer=sec.footer.paragraphs[0];footer.alignment=WD_ALIGN_PARAGRAPH.RIGHT
footer.add_run('研究截至2026年10月25日  ·  ')
fld=OxmlElement('w:fldSimple');fld.set(qn('w:instr'),'PAGE');footer._p.append(fld)
for r in footer.runs:r.font.size=Pt(8)

def add_inline(p,text):
    pattern=r'(\*\*.*?\*\*|\[[^\]]+\]\(https?://[^)]+\))'
    for item in re.split(pattern,text):
        if not item:continue
        if item.startswith('**'):
            p.add_run(item[2:-2]).bold=True
        elif item.startswith('[') and '](' in item:
            label,url=re.match(r'\[([^\]]+)\]\(([^)]+)\)',item).groups()
            hyp=OxmlElement('w:hyperlink');hyp.set(qn('r:id'),p.part.relate_to(url,RT.HYPERLINK,is_external=True))
            r=OxmlElement('w:r');pr=OxmlElement('w:rPr')
            color=OxmlElement('w:color');color.set(qn('w:val'),'245A81');pr.append(color)
            fonts=OxmlElement('w:rFonts')
            for tag in ['ascii','hAnsi','eastAsia']:fonts.set(qn('w:'+tag),'Noto Sans CJK SC')
            pr.append(fonts);r.append(pr);tx=OxmlElement('w:t');tx.text=label;r.append(tx);hyp.append(r);p._p.append(hyp)
        else:p.add_run(item)

def add_table(lines):
    grid=[[v.strip() for v in l.strip('|').split('|')] for l in lines if not re.match(r'^\|[-: |]+\|$',l)]
    n=len(grid[0]); widths={3:[3.2,6.2,7.8],4:[4.1,1.3,6.5,5.3],7:[8.2,1.3,1.3,1.3,1.3,2.0,1.8]}[n]
    t=doc.add_table(rows=0,cols=n);t.autofit=False;t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for col,w in zip(t.columns,widths):col.width=Cm(w)
    for ri,vals in enumerate(grid):
        cells=t.add_row().cells;tr=t.rows[-1]._tr.get_or_add_trPr()
        tr.append(OxmlElement('w:cantSplit'))
        if ri==0:tr.append(OxmlElement('w:tblHeader'))
        for ci,(cell,val) in enumerate(zip(cells,vals)):
            cell.width=Cm(widths[ci]);cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            pr=cell._tc.get_or_add_tcPr(); borders=OxmlElement('w:tcBorders')
            for e in ['top','left','bottom','right']:
                edge=OxmlElement('w:'+e);edge.set(qn('w:val'),'single');edge.set(qn('w:sz'),'4');edge.set(qn('w:color'),'D9D9D9');borders.append(edge)
            pr.append(borders);mar=OxmlElement('w:tcMar')
            for e in ['top','left','bottom','right']:
                el=OxmlElement('w:'+e);el.set(qn('w:w'),'85');el.set(qn('w:type'),'dxa');mar.append(el)
            pr.append(mar);shade=OxmlElement('w:shd');shade.set(qn('w:fill'),'DCE6EE' if ri==0 else ('F5F7F9' if ri%2==0 else 'FFFFFF'));pr.append(shade)
            p=cell.paragraphs[0];p.paragraph_format.space_after=Pt(1);p.paragraph_format.line_spacing=Pt(13)
            p.alignment=WD_ALIGN_PARAGRAPH.CENTER if (n==7 and ci>0) or (n==4 and ci==1) else WD_ALIGN_PARAGRAPH.LEFT
            add_inline(p,val)
            for run in p.runs:run.font.size=Pt(9);run.bold=ri==0
    doc.add_paragraph().paragraph_format.space_after=Pt(0)

lines=md.splitlines();i=0
while i<len(lines):
    line=lines[i]
    if not line.strip():i+=1;continue
    if line.startswith('|'):
        group=[]
        while i<len(lines) and lines[i].startswith('|'):group.append(lines[i]);i+=1
        add_table(group);continue
    if line.startswith('# '):doc.add_paragraph('2025与2026年湖南金种子杯\n金奖项目深度研究',style='Title')
    elif line.startswith('## '):
        p=doc.add_heading(line[3:],1)
        if line.startswith(('## 五','## 六','## 七','## 十')):p.paragraph_format.page_break_before=True
    elif line.startswith('### '):doc.add_heading(line[4:],2)
    elif line.startswith('#### '):doc.add_heading(line[5:],3)
    else:
        p=doc.add_paragraph();add_inline(p,line)
        if line.startswith('依据：'):
            p.paragraph_format.space_after=Pt(9);p.paragraph_format.line_spacing=Pt(12)
            for run in p.runs:run.font.size=Pt(8.5)
        if line.startswith('**官方项目名**'):p.paragraph_format.keep_with_next=True
    i+=1
doc.core_properties.title='2025与2026年湖南金种子杯金奖项目深度研究'
doc.core_properties.subject='36项金奖全样本分类 竞争力分析与创意评分'
doc.core_properties.author=''
docpath=OUT/'2025与2026年金种子杯36项金奖深度研究.docx'
doc.save(docpath)
assert 'turn0' not in md and '' not in md
assert len(re.findall(r'^#### \d\d ',md,re.M))==36
print(json.dumps({'docx':str(docpath),'markdown':str(mdpath),'characters':len(md),'cases':len(data),'sources':len(sources),'categories':dict(collections.Counter(d['category'] for d in data))},ensure_ascii=False,indent=2))
