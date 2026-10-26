from pathlib import Path
from collections import Counter, OrderedDict
import re, json
from urllib.parse import urlparse
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

HERE=Path(__file__).resolve().parent
ROOT=HERE.parents[1]
OUT=ROOT/'outputs/teacher-alumni-20261026'
OUT.mkdir(parents=True,exist_ok=True)
FONT='Noto Sans CJK SC'
LINK=re.compile(r'\[([^\]]+)\]\((https?://[^)]+)\)')

def compose_teacher():
    front=(HERE/'teacher-front.txt').read_text()
    old=(ROOT/'research/candidate-plans-20261025/report-body.txt').read_text()
    part=old.split('## 三 方案 A 道路养护核验')[1].split('## 九 往届金奖')[0]
    part='## 三 方案 A 道路养护核验'+part
    part=re.sub(r'^### (.+)$',r'#### \1',part,flags=re.M)
    part=re.sub(r'^## [三四五六七八] (方案 |备选 )?([A-F] .+)$',r'### \2',part,flags=re.M)
    part=part.replace('学长理“冰峰筑路”','学习长理“冰峰筑路”')
    # The new source establishes pre-provincial-competition product evidence.
    part=part.replace('湖南针对 AI 自习室的治理文件表明','湖南校外培训政策问答表明')
    body=front+'\n\n'+part+'\n\n'+(HERE/'teacher-extended.txt').read_text()+'\n\n'+(HERE/'teacher-plan.txt').read_text()
    body+='\n\n此前两届权重也不同：第二届专业评分为市场 30、创新 25、团队 20、财务 15、社会 10；首届传统产业为 25、25、20、20、10，其余三赛道为 30、25、20、15、10。它们只用于解释当届案例，本届准备仍采用第二章列出的新权重。[第二届通知](https://jyt.hunan.gov.cn/jyt/sjyt/xxgk/tzgg/202510/t20251023_1168604.html) [首届通知](https://www.huhst.edu.cn/cxcy/info/1079/2599.htm)\n'
    refs={}
    for l in (ROOT/'research/gold-study-20261025/sources.txt').read_text().splitlines():
        rid,title,date,url=l.split('|',3);refs[rid]=(title,url)
    cases=json.loads((ROOT/'research/gold-study-20261025/case_analysis.json').read_text())['cases']
    for c in cases:
        if c['id']==1:
            c['evidence_level']='A'
            c['facts']='2025 年 4 月 10 日省教育厅刊载的长理校赛报道已介绍复合改性沥青和视频展示，提供省赛前的产品线索；5 月 26 日学校报道确认传统产业赛道第一。8 月近同名项目介绍只补充技术理解，不倒推全部赛前性能。'
        scores=' / '.join(str(c['scores'][k]) for k in ['I','C','D','V'])
        body+=f"\n### {c['id']:02d} {c['short_name']}\n\n"
        body+=f"{c['year']} 年金奖；{c['official_track']}；{c['schools']}。\n\n"
        body+=f"**正式项目名：** {c['official_name']}\n\n"
        body+=f"**结合与创意评价：** {c['combination']}。I C D V 分项为 {scores}，研究创意分 {c['total']}/100，证据 {c['evidence_level']} 档。\n\n"
        body+=f"**公开证据：** {c['facts']}\n\n"
        body+=f"**竞争力与创意分析：** {c['award_analysis_inference']} {c['creativity_analysis']}\n\n"
        body+=f"**继续核实：** {c['gaps']}\n\n"
        rids=[('R01' if c['year']==2025 else 'R02')]+c['source_ids']
        body+='依据：'+' '.join(f'[{refs[r][0]}]({refs[r][1]})' for r in dict.fromkeys(rids))
        if c['id']==1:body+=' [冰峰筑路校赛赛前报道](https://jyt.hunan.gov.cn/jyt/sjyt/gxdt/202504/t20250410_33637313.html)'
        body+='\n'
    return body

def setfont(style,size,bold=False):
    style.font.name=FONT;style.font.size=Pt(size);style.font.bold=bold;style.font.color.rgb=RGBColor(0,0,0)
    rpr=style._element.get_or_add_rPr();f=rpr.get_or_add_rFonts()
    for k in list(f.attrib):
        if 'theme' in k.lower():del f.attrib[k]
    for k in ['ascii','hAnsi','eastAsia','cs']:f.set(qn('w:'+k),FONT)
    for b in style._element.xpath('./w:pPr/w:pBdr'):b.getparent().remove(b)

def hyperlink(p,label,url=None,anchor=None,size=11):
    h=OxmlElement('w:hyperlink')
    if url:h.set(qn('r:id'),p.part.relate_to(url,RT.HYPERLINK,is_external=True))
    if anchor:h.set(qn('w:anchor'),anchor)
    r=OxmlElement('w:r');pr=OxmlElement('w:rPr')
    f=OxmlElement('w:rFonts')
    for k in ['ascii','hAnsi','eastAsia']:f.set(qn('w:'+k),FONT)
    pr.append(f)
    for tag,val in [('color','000000'),('sz',str(int(size*2))),('u','single')]:
        e=OxmlElement('w:'+tag);e.set(qn('w:val'),val);pr.append(e)
    r.append(pr);t=OxmlElement('w:t');t.text=label;r.append(t);h.append(r);p._p.append(h)

def inline(p,text,sources,size=None):
    for frag in re.split(r'(\*\*.*?\*\*|\[[^\]]+\]\(https?://[^)]+\))',text):
        if not frag:continue
        m=LINK.fullmatch(frag)
        if m:
            label,url=m.groups();n=list(sources).index(url)+1
            hyperlink(p,f'[S{n:03d}]',url,size=size or 11)
        else:
            b=frag.startswith('**');r=p.add_run(frag[2:-2] if b else frag);r.bold=b
            if size:r.font.size=Pt(size)

def table(doc,lines,sources):
    grid=[[v.strip() for v in l.strip('|').split('|')] for l in lines if not ('-' in l and re.match(r'^\|[-: |]+\|$',l))]
    n=len(grid[0]);assert all(len(r)==n for r in grid)
    total=17.59
    widths={2:[5,12.59],3:[3.7,6.7,7.19],4:[3.3,4.3,4.6,5.39],5:[6.39,2.8,2.8,2.8,2.8],7:[4.39,2.2,2.2,2.2,2.2,2.2,2.2]}.get(n,[total/n]*n)
    a=grid[0]
    if a[0]=='专业评审项目' or a[0]=='大众评审项目':widths=[5,1.6,10.99]
    if a[0]=='本届官方赛道':widths=[5.8,11.79]
    if a[0]=='编号与方向':widths=[4.7,6.3,6.59]
    if a[0]=='方向' and a[1]=='控制额':widths=[2.3,2.3,6.4,6.59]
    if a[0]=='时间':widths=[4.0,6.6,6.99]
    if a[0]=='历史聊天':widths=[4.7,8.29,4.6]
    if a[0]=='问题编号':widths=[2.0,6.2,4.2,5.19]
    tb=doc.add_table(rows=0,cols=n);tb.autofit=False;tb.alignment=WD_TABLE_ALIGNMENT.CENTER
    for col,w in zip(tb.columns,widths):col.width=Cm(w)
    for ri,vals in enumerate(grid):
        row=tb.add_row();tr=row._tr.get_or_add_trPr();tr.append(OxmlElement('w:cantSplit'))
        if ri==0:tr.append(OxmlElement('w:tblHeader'))
        for ci,(cell,val,w) in enumerate(zip(row.cells,vals,widths)):
            cell.width=Cm(w);cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            tc=cell._tc.get_or_add_tcPr();b=OxmlElement('w:tcBorders')
            for e in ['top','bottom','left','right']:
                x=OxmlElement('w:'+e)
                for k,v in [('val','single'),('sz','4'),('color','D9D9D9')]:x.set(qn('w:'+k),v)
                b.append(x)
            tc.append(b);shade=OxmlElement('w:shd');shade.set(qn('w:fill'),'E6EDF3' if ri==0 else 'FFFFFF');tc.append(shade)
            mar=OxmlElement('w:tcMar')
            for e in ['top','bottom','left','right']:
                x=OxmlElement('w:'+e);x.set(qn('w:w'),'85');x.set(qn('w:type'),'dxa');mar.append(x)
            tc.append(mar)
            p=cell.paragraphs[0];p.paragraph_format.space_after=Pt(2);p.paragraph_format.line_spacing=Pt(15)
            if ri==0 or (n>=5 and len(grid)<=8 and ri<len(grid)-1):p.paragraph_format.keep_with_next=True
            inline(p,val,sources,10.5)
            if n>=5 and ci>0:p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            if a[0]=='方向' and a[1]=='控制额' and ci==1:p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:r.bold=ri==0
            if not any(vals):
                p.add_run(' ');p.paragraph_format.space_after=Pt(17)
    p=doc.add_paragraph();p.paragraph_format.space_after=Pt(0);p.paragraph_format.line_spacing=Pt(4)

def bookmark(p,name,i):
    s=OxmlElement('w:bookmarkStart');s.set(qn('w:id'),str(i));s.set(qn('w:name'),name)
    e=OxmlElement('w:bookmarkEnd');e.set(qn('w:id'),str(i));p._p.insert(0,s);p._p.append(e)

def build(body,filename,teacher=False):
    sources=OrderedDict()
    for label,url in LINK.findall(body):sources.setdefault(url,label)
    doc=Document();sec=doc.sections[0];sec.page_width=Cm(21.59);sec.page_height=Cm(27.94)
    sec.top_margin=Cm(1.8);sec.bottom_margin=Cm(1.8);sec.left_margin=Cm(2);sec.right_margin=Cm(2);sec.footer_distance=Cm(.7)
    for name,size,b in [('Normal',11,False),('Title',23,True),('Heading 1',16,True),('Heading 2',13,True),('Heading 3',11.5,True)]:setfont(doc.styles[name],size,b)
    p=doc.styles['Normal'].paragraph_format;p.line_spacing=Pt(17);p.space_after=Pt(7);p.widow_control=True
    doc.styles['Title'].paragraph_format.space_after=Pt(18)
    for n in [1,2,3]:
        p=doc.styles[f'Heading {n}'].paragraph_format;p.keep_with_next=True;p.space_before=Pt(11 if n>1 else 0);p.space_after=Pt(7)
    footer=sec.footer.paragraphs[0];footer.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    for instr in ['PAGE','NUMPAGES']:
        if instr=='NUMPAGES':footer.add_run(' / ')
        fld=OxmlElement('w:fldSimple');fld.set(qn('w:instr'),instr)
        r=OxmlElement('w:r');t=OxmlElement('w:t');t.text='1';r.append(t);fld.append(r);footer._p.append(fld)
    headings=re.findall(r'^## (.+)$',body,re.M)
    bibliography='十一 公开资料来源' if teacher else '九 来源索引'
    headings.append(bibliography)
    tocpath=HERE/'teacher-toc.json';toc=json.loads(tocpath.read_text()) if tocpath.exists() else {}
    inserted=False;hidx=0;lines=body.splitlines();i=0
    while i<len(lines):
        line=lines[i].strip()
        if not line:i+=1;continue
        if line.startswith('|'):
            g=[]
            while i<len(lines) and lines[i].strip().startswith('|'):g.append(lines[i].strip());i+=1
            table(doc,g,sources);continue
        if line.startswith('# '):doc.add_paragraph(line[2:],style='Title')
        elif line.startswith('## '):
            if teacher and not inserted:
                doc.add_page_break();p=doc.add_paragraph('目录',style='Heading 1')
                for j,h in enumerate(headings,1):
                    p=doc.add_paragraph();p.paragraph_format.space_after=Pt(12)
                    p.paragraph_format.tab_stops.add_tab_stop(Cm(17.0),WD_TAB_ALIGNMENT.RIGHT,WD_TAB_LEADER.DOTS)
                    hyperlink(p,h,anchor=f'chapter_{j}')
                    p.add_run('\t'+str(toc.get(h,'—')))
                doc.add_paragraph('现场交流可先阅读第一章、第三章和第八章；具体方案在第四至六章，全部金奖分析在第十章。')
                inserted=True
            hidx+=1;p=doc.add_heading(line[3:],1);bookmark(p,f'chapter_{hidx}',hidx)
            if (teacher and hidx in [1,4,10]) or (not teacher and hidx in [2,3,8]):p.paragraph_format.page_break_before=True
            else:p.paragraph_format.space_before=Pt(17)
        elif line.startswith('### '):doc.add_heading(line[4:],2)
        elif line.startswith('#### '):doc.add_heading(line[5:],3)
        else:
            p=doc.add_paragraph();inline(p,line,sources)
            following=next((l.strip() for l in lines[i+1:] if l.strip()),'')
            if teacher and hidx==10 and (re.match(r'^20\d\d 年金奖',line) or line.startswith('**正式项目名') or line.startswith('**结合与创意评价')):
                p.paragraph_format.keep_with_next=True
            if not teacher and hidx>=3 and hidx<=6 and following and not following.startswith('#'):
                p.paragraph_format.keep_with_next=True
        i+=1
    p=doc.add_heading(bibliography,1);bookmark(p,f'chapter_{len(headings)}',len(headings));p.paragraph_format.page_break_before=True
    doc.add_paragraph('正文中的来源编号可点击打开原始页面，以下保留对应标题和网站。正式名单用于确认奖项；产品页支持公开能力描述；研究论文与案例报道按正文注明的时间和适用范围使用。')
    for j,(url,label) in enumerate(sources.items(),1):
        p=doc.add_paragraph();p.paragraph_format.line_spacing=Pt(15);p.paragraph_format.space_after=Pt(5)
        p.add_run(f'[S{j:03d}] ');hyperlink(p,label,url,size=10.5)
        r=p.add_run('  '+urlparse(url).netloc);r.font.size=Pt(10.5)
    doc.core_properties.title=lines[0][2:];doc.core_properties.subject='金种子杯研究与选题' if teacher else '过往获奖选手经验咨询'
    doc.core_properties.author='';doc.core_properties.last_modified_by=''
    path=OUT/filename;doc.save(path)
    (HERE/(('teacher' if teacher else 'alumni')+'-source-index.json')).write_text(json.dumps([{'id':f'S{j:03d}','title':label,'url':url} for j,(url,label) in enumerate(sources.items(),1)],ensure_ascii=False,indent=2))
    return {'path':str(path),'characters':len(body),'sources':len(sources),'chapters':headings,'tables':len(doc.tables),'paragraphs':len(doc.paragraphs)}

teacher=compose_teacher();alumni=(HERE/'alumni.txt').read_text()
assert all(f'### Q{i:02d} ' in alumni for i in range(1,35))
assert not re.search(r'|turn\d+(?:search|view)',teacher+alumni)
(HERE/'teacher-complete.txt').write_text(teacher)
results=[build(teacher,'01_第三届金种子杯调研方案_老师汇报.docx',True),build(alumni,'02_金种子杯获奖学长咨询疑问清单.docx')]
(HERE/'build-manifest.json').write_text(json.dumps(results,ensure_ascii=False,indent=2))
print(json.dumps(results,ensure_ascii=False,indent=2))
