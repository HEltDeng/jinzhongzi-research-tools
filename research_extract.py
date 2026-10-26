from pathlib import Path
import pdfplumber,json,re,collections
BASE=Path(__file__).resolve().parent
RAW=BASE/"research"/"sources"
tracks25=["传统产业赛道","优势产业赛道","新兴产业赛道","未来产业赛道"]
tracks26=["先进制造与高端装备赛道","人工智能与信息技术赛道","新能源与新材料赛道","空天海洋与低空经济赛道","现代农业与大健康产业赛道","新消费与AI+文化科技赛道","AI 3D大模型应用专项赛"]
def clean(s):return (s or "").replace("\n","").strip()
records=[]
def read_pdf(path,year,track=None,url=""):
 doc=pdfplumber.open(path)
 alltext=[]
 for page_num,p in enumerate(doc.pages,1):
  text=p.extract_text() or ""
  alltext.append(text)
  if year==2025:
   for t in tracks25:
    if t+"获奖名单" in re.sub(r"\s","",text):track=t
  if not track:continue
  for table in p.extract_tables():
   for row in table:
    if len(row)<6:continue
    row=[clean(x) for x in row]
    if row[0].isdigit() and row[2] in ["金奖","银奖","铜奖","优胜奖"]:
     records.append({"届次":1 if year==2025 else 2,"获奖年份":year,"官方赛道":track,"赛道序号":int(row[0]),"项目名称":row[1],"奖项":row[2],"名单首位队员":re.split("[、,，]",row[3])[0].strip(),"所属学校":row[5],"名单页码":page_num,"原始附件":str(path.relative_to(BASE)),"名单来源":url})
    elif row[0]=="" and any(row[1:]) and records:
     # Continued cells across page breaks are reviewed in extracted output.
     print("CONTINUATION",path.name,page_num,row)
 (path.with_suffix(".txt")).write_text("\n\n".join(alltext),encoding="utf-8")
read_pdf(RAW/"2025_awards_final_attachment1.pdf",2025,url="https://hunan.gov.cn/topic/dxscy/dxscyZCJC/dxscysjzc/202510/t20251015_33826774.html")
for i,t in enumerate(tracks26):
 read_pdf(RAW/f"2026_awards_final_attachment{2*i+1}.pdf",2026,t,"https://cms.jyt.hunan.gov.cn/zcms/contentcore/resource/download?ID="+str(131085+i))
(BASE/"research"/"awards.json").write_text(json.dumps(records,ensure_ascii=False,indent=2),encoding="utf-8")
print("COUNTS",collections.Counter((r["获奖年份"],r["官方赛道"],r["奖项"]) for r in records))
print("TOTAL",len(records))
for r in records:
 if r["奖项"]=="金奖" or "长沙理工" in r["所属学校"]:
  print(r["获奖年份"],r["官方赛道"],r["奖项"],r["项目名称"],r["所属学校"])
