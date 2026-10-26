import fs from "node:fs/promises";
import path from "node:path";
import {fileURLToPath} from "node:url";
import {Workbook, SpreadsheetFile} from "@oai/artifact-tool";

const base=path.dirname(fileURLToPath(import.meta.url));
const out=path.join(base,"outputs/jzz-research-20261020");
const qa=path.join(base,"research/qa/workbook");
await fs.mkdir(out,{recursive:true});await fs.mkdir(qa,{recursive:true});
const read=async name=>JSON.parse(await fs.readFile(path.join(base,"research",name),"utf8"));
const awards=await read("awards_enriched.json"), gold=await read("gold_analysis.json"), sources=await read("sources_index.json");
const wb=Workbook.create();
const summary=wb.worksheets.add("赛道统计");
const all=wb.worksheets.add("获奖全表");
const golden=wb.worksheets.add("36项金奖拆解");
const csust=wb.worksheets.add("长理相关");
const refs=wb.worksheets.add("参考来源");
const C=n=>{let s="";while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s};
const title=(sh,text,note)=>{
  sh.showGridLines=false;
  sh.getRange("A2").values=[[text]];
  sh.getRange("A2").format.font={name:"Arial",size:16,bold:true,color:"#162A3D"};
  sh.getRange("A3").values=[[note]];
  sh.getRange("A3").format.font={name:"Arial",size:10,italic:true,color:"#666666"};
  sh.getRange("A2:H2").format.rowHeight=25;
};
function tab(sh,headers,rows,widths,name){
  const end=5+rows.length, last=C(headers.length);
  sh.getRange(`A5:${last}${end}`).values=[headers,...rows];
  const t=sh.tables.add(`A5:${last}${end}`,true,name);t.showFilterButton=true;
  const r=sh.getRange(`A5:${last}${end}`);
  r.format.font={name:"Arial",size:10,color:"#263746"};
  r.format.verticalAlignment="center";
  r.format.wrapText=true;
  sh.getRange(`A5:${last}5`).format.fill="#233C57";
  sh.getRange(`A5:${last}5`).format.font={name:"Arial",size:10,bold:true,color:"#FFFFFF"};
  sh.getRange(`A5:${last}5`).format.horizontalAlignment="center";
  sh.getRange(`A5:${last}5`).format.rowHeight=32;
  widths.forEach((w,i)=>sh.getRange(`${C(i+1)}5:${C(i+1)}${end}`).format.columnWidthPx=w);
  rows.forEach((row,i)=>{
    let lineCount=1;
    row.forEach((v,j)=>{
      const len=[...String(v??"")].reduce((n,ch)=>n+(/[^\x00-\xff]/.test(ch)?1.85:1),0);
      lineCount=Math.max(lineCount,Math.ceil(len/((widths[j]-18)/6.5)));
    });
    sh.getRange(`A${i+6}:${last}${i+6}`).format.rowHeight=Math.max(34,lineCount*13+10);
  });
  sh.freezePanes.freezeRows(5);
  return end;
}
title(summary,"金种子杯 · 两届最终获奖分布","研究截至2026-10-20；统计官方获奖条目，不代表全部参赛项目或获奖率。");
summary.tabColor="#233C57";
title(all,"364项官方获奖项目","官方字段来自最终附件；主题标签仅按题名初标，不能证明实际采用某项技术。");
title(golden,"36项金奖 · 交叉方向与证据","政策关联与竞争力分析均为研究推断；不代表评委意见。缺少证据的项目明确标为待核。");
golden.tabColor="#456882";
title(csust,"长理相关 · 包含联合学校条目","按所属学校字段含“长沙理工”筛选：2025年15项，2026年12项；不等同校方牵头统计口径。");
title(refs,"参考来源与适用范围","官方名单优先；自述、赛后资料和政策时间在记录中区分。URL可复制到浏览器。");

const ah=["获奖年份","官方赛道","奖项","项目名称（名单提取）","所属学校","参赛队员首位","赛道序号","主题标签（题名初标）","长理相关","附件页码","名单来源URL","本地原始附件"];
const av=awards.map(r=>[r["获奖年份"],r["官方赛道"],r["奖项"],r["项目名称"],r["所属学校"],r["名单首位队员"],r["赛道序号"],r["主题标签_题名初标"],r["长理相关"],r["名单页码"],r["名单来源"],r["原始附件"]]);
tab(all,ah,av,[85,215,75,410,310,110,80,250,80,80,440,310],"AwardRecords");
all.freezePanes.freezeColumns(4);
all.getRange("A6:A369").setNumberFormat("0");
all.getRange("G6:G369").setNumberFormat("0");
all.getRange("J6:J369").setNumberFormat("0");

const tracks=[...new Set(awards.map(r=>`${r["获奖年份"]}|${r["官方赛道"]}`))].map(s=>s.split("|"));
const sv=tracks.map(([year,track])=>[Number(year),track,null,null,null,null,null,Number(year)===2025?"S01":"S02"]);
tab(summary,["年份","官方赛道","金奖","银奖","铜奖","优胜奖","合计","来源"],sv,[80,310,80,80,80,90,80,80],"TrackCounts");
summary.freezePanes.unfreeze();
for(let i=0;i<tracks.length;i++){
  const row=i+6;
  summary.getRange(`C${row}:F${row}`).formulas=[["金奖","银奖","铜奖","优胜奖"].map((_,j)=>`=COUNTIFS('获奖全表'!$A$6:$A$369,$A${row},'获奖全表'!$B$6:$B$369,$B${row},'获奖全表'!$C$6:$C$369,${C(j+3)}$5)`)];
  summary.getRange(`G${row}`).formulas=[[`=SUM(C${row}:F${row})`]];
}
summary.getRange("A19:G22").values=[
["年度汇总",null,"金奖","银奖","铜奖","优胜奖","合计"],
[2025,null,null,null,null,null,null],
[2026,null,null,null,null,null,null],
["两届合计",null,null,null,null,null,null]
];
summary.getRange("A19:G22").format.font={name:"Arial",size:10};
summary.getRange("A19:G19").format.fill="#E9EFF4";
summary.getRange("A19:G19").format.font.bold=true;
summary.getRange("A19:G22").format.rowHeight=28;
for(const row of [20,21]){
  summary.getRange(`C${row}:F${row}`).formulas=[Array.from({length:4},(_,j)=>`=COUNTIFS('获奖全表'!$A$6:$A$369,$A${row},'获奖全表'!$C$6:$C$369,${C(j+3)}$19)`)];
  summary.getRange(`G${row}`).formulas=[[`=SUM(C${row}:F${row})`]];
}
summary.getRange("C22:G22").formulas=[Array.from({length:5},(_,j)=>`=SUM(${C(j+3)}20:${C(j+3)}21)`)];
summary.getRange("A22:G22").format.font.bold=true;
summary.getRange("A22:G22").format.borders={bottom:{style:"thin",color:"#233C57"}};
summary.getRange("C6:G22").setNumberFormat("0");
const notes=[
"范围：首届2025年、第二届2026年；未取得全部报名者及评委逐项评分。",
"口径：2026年早期公示206项；7月正式附件205项，本库采用正式版。",
"解释：赛道名额、获奖数量不能用于比较选题获奖概率。",
"当前：第三届2026-12-31报名截止，2027年5月上旬决赛；长理校内时间待核。",
"用法：全表可按年份、赛道、奖项、学校筛选；金奖拆解中的S编号见参考来源。",
"更新：本文件为两届名单快照；新增记录需重新生成统计及长理关联视图。",
"建议：先验证服务区充电预测与容量调度；养护复发识别与验收助手为备选。"
];
summary.getRange("A25:A31").values=notes.map(x=>[x]);
summary.getRange("A25:A31").format.font={name:"Arial",size:10,color:"#555555"};
summary.getRange("A25:A31").format.rowHeight=23;
summary.getRange("A25:A31").format.wrapText=false;

const gh=["年份","官方赛道","金奖项目名称","交叉方向（分析）","技术/产品解释","同期政策编号","竞争力假设（非获奖评语）","证据等级","补充来源编号","待核实事项","名单来源URL","所属学校"];
const gv=gold.map(r=>[r["获奖年份"],r["官方赛道"],r["项目名称"],r["交叉方向"],r["技术产品解释"],r["同期政策编号"],r["竞争力假设_非官方获奖理由"],r["证据等级"],r["案例补充来源编号"]||"未取得补充材料",r["待核实事项"],r["名单来源"],r["所属学校"]]);
tab(golden,gh,gv,[75,210,390,230,270,130,440,230,150,370,440,270],"GoldAnalysis");
golden.freezePanes.freezeColumns(3);
golden.getRange("A6:A41").setNumberFormat("0");
const ci=awards.map((r,i)=>r["长理相关"]==="是"?i:-1).filter(i=>i>=0);
tab(csust,ah,ci.map(i=>av[i]),[85,215,75,410,310,110,80,250,80,80,440,310],"CsustRecords");
csust.getRange("A6:L32").formulas=ci.map(i=>Array.from({length:12},(_,j)=>`='获奖全表'!${C(j+1)}${i+6}`));
csust.getRange("A6:A32").setNumberFormat("0");
csust.freezePanes.freezeColumns(4);
tab(refs,["编号","资料标题","发布/时间口径","来源性质","本研究用途与限制","原始URL"],sources.map(s=>[s.id,s.title,s.date,s.type,s.use,s.url]),[75,370,230,170,470,500],"SourceRegistry");
refs.freezePanes.freezeColumns(2);

wb.recalculate();
const inspect=await wb.inspect({kind:"region",sheetId:"赛道统计",range:"A19:G22",maxChars:3000,tableMaxRows:5,tableMaxCols:7});
console.log(inspect.ndjson??inspect);
await fs.writeFile(path.join(qa,"summary_values.json"),JSON.stringify(summary.getRange("A5:H22").values,null,2));
const previews=[
["赛道统计","A2:H31","summary"],
["获奖全表","A2:F10","all_start"],
["获奖全表","A362:F369","all_end"],
["36项金奖拆解","A2:E10","gold_start"],
["36项金奖拆解","F5:J10","gold_evidence"],
["长理相关","A2:F10","csust"],
["参考来源","A2:E10","sources"]
];
for(const [sheetName,range,name] of previews){
  const image=await wb.render({sheetName,range,format:"png",scale:1.3});
  await fs.writeFile(path.join(qa,name+".png"),new Uint8Array(await image.arrayBuffer()));
}
const xlsx=await SpreadsheetFile.exportXlsx(wb);
const dest=path.join(out,"金种子杯364项获奖项目与36项金奖分析.xlsx");
await xlsx.save(dest);
console.log(dest);
