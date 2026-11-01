from pathlib import Path
import json, re, concurrent.futures, subprocess
from lxml import html
from urllib.parse import urljoin

BASE=Path(__file__).resolve().parent
RAW=BASE/"research"/"sources"
RAW.mkdir(parents=True,exist_ok=True)
SOURCES={
 "2025_awards_notice":"https://jyt.hunan.gov.cn/jyt/sjyt/xxgk/tzgg/202505/t20250526_1160367.html",
 "2025_awards_final":"https://hunan.gov.cn/topic/dxscy/dxscyZCJC/dxscysjzc/202510/t20251015_33826774.html",
 "2026_awards_final":"https://jyt.hunan.gov.cn/jyt/sjyt/xxgk/tzgg/202607/t20260708_1180348.html",
 "2026_rules":"https://jyt.hunan.gov.cn/jyt/sjyt/xxgk/tzgg/202510/t20251023_1168604.html",
 "2027_launch":"https://www.hunan.gov.cn/hnszf/hnyw/zwdt/202609/t20260930_34074643.html",
 "2025_rules":"https://www.huhst.edu.cn/cxcy/info/1079/2599.htm",
 "2025_work_report":"https://www.hunan.gov.cn/hnszf/hnyw/sy/hnyw1/202501/t20250127_33576548.html",
 "2026_work_report":"https://www.hunan.gov.cn/hnszf/hnyw/sy/hnyw1/202602/t20260224_33919204.html",
 "startup_policy":"https://www.hunan.gov.cn/hnszf/xxgk/wjk/szfbgt/202411/t20241129_33514183.html",
 "2027_rules_repost":"https://www.changdecollege.cn/rwysxy/info/1081/6493.htm",
}

def get(url):
 p=subprocess.run(["curl","-L","--fail","--max-time","45","-sS",url],capture_output=True)
 if p.returncode: raise RuntimeError(p.stderr.decode()[:300])
 return p.stdout

def fetch(item):
 key,url=item
 data=get(url)
 text=data.decode("utf-8",errors="replace")
 (RAW/f"{key}.html").write_text(text,encoding="utf-8")
 soup=html.fromstring(text)
 (RAW/f"{key}.txt").write_text(soup.text_content(),encoding="utf-8")
 links=[]
 for a in soup.xpath("//a[@href]"):
  href=urljoin(url,a.get("href"))
  if re.search(r"\.(pdf|xlsx|docx|doc|xls)(?:\?|$)|resource/download|wbfile/.*download",href,re.I):
   links.append({"name":a.text_content().strip(),"url":href})
 result={"id":key,"url":url,"attachments":links,"status":200}
 print(json.dumps(result,ensure_ascii=False),flush=True)
 for n,l in enumerate(links):
  match=re.search(r"\.(pdf|xlsx|docx|doc|xls)(?:\?|$)",l["url"],re.I)
  ext=match.group(1) if match else ("xlsx" if key=="2025_awards_notice" else "pdf" if "awards" in key or "2027" in key else "doc")
  p=RAW/f"{key}_attachment{n+1}.{ext}"
  try:
   b=get(l["url"]);p.write_bytes(b)
   l["local"]=str(p.relative_to(BASE))
  except Exception as e: l["error"]=str(e)
 return result

if __name__=="__main__":
 results=[]
 with concurrent.futures.ThreadPoolExecutor(max_workers=6) as pool:
  futs={pool.submit(fetch,item):item for item in SOURCES.items()}
  for fut in concurrent.futures.as_completed(futs):
   try: results.append(fut.result())
   except Exception as e: results.append({"id":futs[fut][0],"error":str(e)})
 (RAW/"manifest.json").write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding="utf-8")
