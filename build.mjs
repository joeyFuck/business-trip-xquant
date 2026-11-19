import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from 'file:///C:/Program%20Files/WindowsApps/OpenAI.CodexPrimaryRuntime.v26-1112-611-0_26.1112.611.0_x64__3k8sg7r9htsxt/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';
const raw=JSON.parse(await fs.readFile('oa_rows.json','utf8'));
if(raw.length!==81)throw Error('条数不符');
const rows=raw.map(([title,type,date,status])=>{title=title.replace(/ 新到达流程$/,''); const m=title.match(/(?:^| )(\d+(?:\.\d+)?) (?=出差申请及报销流程|差旅费报销流程)/);return [Number(date.slice(0,4)),new Date(date.replace(' ','T')+'+08:00').getTime()/86400000+25569+8/24,type==='差旅费报销流程'?'财务管理':'人力资源',type,title.split(/\d{4}-\d{2}-\d{2} /)[1]?.replace(/^\S+ /,'')||'',m?Number(m[1]):null,status,title];}).sort((a,b)=>a[1]-b[1]);
const expenses=JSON.parse(await fs.readFile('expense_rows.json','utf8'));
const normalize=t=>t.replace(/ 新到达流程$/,'').replace(/\s+/g,' ').trim();
const byTitle=new Map();for(const e of expenses){const k=normalize(e.流程标题);if(!byTitle.has(k))byTitle.set(k,[]);byTitle.get(k).push(e);}
const amount=e=>{const a=e.明细类型==='携程公司垫付'?e.结算金额:(e['财务核定金额（预算扣减金额）']!==undefined&&e['财务核定金额（预算扣减金额）']!==''?e['财务核定金额（预算扣减金额）']:e.核定金额!==undefined&&e.核定金额!==''?e.核定金额:e.报销金额);if(a==null||a===''||!Number.isFinite(Number(a)))throw Error('金额缺失');return Math.round(Number(a)*100);};
const categories=[/飞机|火车|市内交通费/,/住宿费/,/出差补贴/];
const missing=[];for(const row of rows){const ee=byTitle.get(normalize(row[7]));if(!ee){row.push('暂无明细','暂无明细','暂无明细');missing.push(row[7]);continue;}row.push(...categories.map(re=>ee.filter(e=>re.test(e.费用项目)).reduce((a,e)=>a+amount(e),0)/100));}
for(let k=0;k<3;k++){const expected=expenses.filter(e=>categories[k].test(e.费用项目)).reduce((a,e)=>a+amount(e),0);const actual=rows.reduce((a,r)=>a+(typeof r[8+k]==='number'?Math.round(r[8+k]*100):0),0);if(actual!==expected)throw Error('分类合计不符');console.log(['交通费','住宿费','出差补贴'][k]+': '+actual/100);}
if(missing.length!==4)throw Error('流程匹配异常:'+JSON.stringify(missing));console.log(JSON.stringify({matched:81-missing.length,missing,total:rows.reduce((a,r)=>a+(typeof r[8]==='number'?r[8]:0),0)}));
const wb=Workbook.create(); const sum=wb.worksheets.add('年度汇总'), d=wb.worksheets.add('流程明细');
d.getRange('A1').values=[['来源：公司 OA 我的请求；两类流程全部记录；读取日期 2026-11-18']];
d.getRange('A2').values=[['https://oa.xquant.com/wui/index.html#/main/workflow/listMine?menuIds=1,14&menuPathIds=1,14&_key=jwo4nq']];
d.getRange('A4:K4').values=[['创建年度','流程创建时间','来源类别','流程名称','事由（标题提取）','标题金额（元）','当前节点','原始流程标题','交通费（元）','住宿费（元）','出差补贴（元）']];d.getRange('A5:K85').values=rows;d.tables.add('A4:K85',true,'TravelRecords');d.freezePanes.freezeRows(4);d.getRange('B5:B85').setNumberFormat('yyyy-mm-dd hh:mm:ss');d.getRange('F5:F85').setNumberFormat('#,##0.00');
sum.getRange('A2').values=[['2019—2026 出差流程汇总']];sum.getRange('A4:G4').values=[['创建年度','流程总数','已归档','行程取消','其他在途','标题金额合计（元）','已归档标题金额（元）']];
for(let i=0;i<8;i++){let r=i+5;sum.getRange(`A${r}`).values=[[2019+i]];sum.getRange(`B${r}:G${r}`).formulas=[[`=COUNTIF('流程明细'!A5:A85,A${r})`,`=COUNTIFS('流程明细'!A5:A85,A${r},'流程明细'!G5:G85,"归档")`,`=COUNTIFS('流程明细'!A5:A85,A${r},'流程明细'!G5:G85,"3-行程取消")`,`=B${r}-C${r}-D${r}`,`=SUMIF('流程明细'!A5:A85,A${r},'流程明细'!F5:F85)`,`=SUMIFS('流程明细'!F5:F85,'流程明细'!A5:A85,A${r},'流程明细'!G5:G85,"归档")`]];}
sum.getRange('A13').values=[['合计']];sum.getRange('B13:G13').formulas=[['=SUM(B5:B12)','=SUM(C5:C12)','=SUM(D5:D12)','=SUM(E5:E12)','=SUM(F5:F12)','=SUM(G5:G12)']];sum.getRange('F5:G13').setNumberFormat('#,##0.00');
sum.getRange('A16').values=[['口径：按流程创建年度统计，流程条数不等同实际出差次数。']];sum.getRange('A17').values=[['金额来自列表标题，不代表已支付金额；未填写金额保留空白。']];sum.getRange('A18').values=[['包含旧流程18条、新流程63条；取消及在途记录保留并分列统计。']];sum.getRange('A19').values=[['未读取表单内的出发/返回日期、出差天数及费用明细。统计截至2026-11-18。']];
for(const sh of [sum,d]){sh.showGridLines=false;sh.getUsedRange().format.font={name:'Arial',size:11};sh.getUsedRange().format.rowHeight=25;sh.getRange(sh===sum?'A4:G4':'A4:H4').format={fill:'#23466A',font:{color:'#FFFFFF',bold:true},rowHeight:36,wrapText:true};}
sum.getRange('A1:G19').format.columnWidth=18;sum.getRange('F1:G19').format.columnWidth=27;sum.getRange('A2').format.font={size:16,bold:true};
for(const [col,width] of [['A',12],['B',24],['C',14],['D',26],['E',55],['F',20],['G',20],['H',105]]) d.getRange(`${col}4:${col}85`).format.columnWidth=width;
d.getRange('E5:E85').format.wrapText=true;d.getRange('H5:H85').format.wrapText=true;d.getRange('A5:H85').format.rowHeight=48;
sum.getRange('H4:J4').values=[['交通费（元）','住宿费（元）','出差补贴（元）']];for(const [dest,src] of [['H','I'],['I','J'],['J','K']]){for(let r=5;r<=12;r++)sum.getRange(`${dest}${r}`).formulas=[[`=SUMIF('流程明细'!$A$5:$A$85,A${r},'流程明细'!$${src}$5:$${src}$85)`]];sum.getRange(`${dest}13`).formulas=[[`=SUM(${dest}5:${dest}12)`]];}sum.getRange('H5:J13').setNumberFormat('#,##0.00');d.getRange('I5:K85').setNumberFormat('#,##0.00');
sum.getRange('A17').values=[['原金额列来自列表标题，不代表已支付金额；未填写金额保留空白。']];
sum.getRange('A19').values=[['交通费含飞机、火车、市内交通费；住宿费及出差补贴按同名费用项目统计，含公司垫付。']];
sum.getRange('A20').values=[['优先用财务核定/核定金额，未核定用报销金额；公司垫付用结算金额。其他、预订服务费不混入三列。']];
sum.getRange('A21').values=[['77条有明细；2条尚无报销信息、2条取消未提取，标“暂无明细”；年度为已知金额小计。']];
for(const [sh,range] of [[sum,'A4:J4'],[d,'A4:K4']])sh.getRange(range).format={fill:'#23466A',font:{color:'#FFFFFF',bold:true},rowHeight:48,wrapText:true};sum.getRange('H1:J21').format.columnWidth=20;d.getRange('I4:K85').format.columnWidth=20;sum.getRange('A20:J21').format.font={name:'Arial',size:11};
wb.recalculate();console.log((await wb.inspect({kind:'table',range:'年度汇总!A4:H13',tableMaxRows:10,tableMaxCols:8})).ndjson);console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?',options:{useRegex:true,maxResults:10}})).ndjson);
console.log((await wb.inspect({kind:'table',range:'年度汇总!H4:J13',tableMaxRows:10,tableMaxCols:3})).ndjson);
await fs.mkdir('outputs/oa_summary',{recursive:true});for(const [sheet,range,name] of [['年度汇总','F4:J13','summary'],['流程明细','I4:K10','detail']]){let p=await wb.render({sheetName:sheet,range,scale:1});await fs.writeFile(`outputs/oa_summary/${name}.png`,new Uint8Array(await p.arrayBuffer()));}
await (await SpreadsheetFile.exportXlsx(wb)).save('outputs/oa_summary/出差流程汇总_2019至2026_费用分类版.xlsx');
