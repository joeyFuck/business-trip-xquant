import fs from 'node:fs/promises';
const dir='outputs/oa_details';const records=[...JSON.parse(await fs.readFile(dir+'/old_manifest.json')),...JSON.parse(await fs.readFile(dir+'/new_manifest.json'))];
let output=[],checks=[];
for(const rec of records){let s=await fs.readFile(dir+'/verified/'+rec.id+'.ui.txt','utf8');let lines=s.split('\n');let tables=[];for(let i=0;i<lines.length;i++){if(!/table \(表格\)/.test(lines[i]))continue;let indent=lines[i].match(/^\s*/)[0].length;let end=i+1;while(end<lines.length&&lines[end].match(/^\s*/)[0].length>indent)end++;let block=lines.slice(i+1,end).join('\n');if(/table \(表格\)/.test(block))continue;let rows=block.split(/\d+ data item \(行\)/).map(b=>[...b.matchAll(/data item \(项\)(?: ([^\n]*))?/g)].map(m=>m[1]||''));let hi=rows.findIndex(r=>r.includes('费用项目'));if(hi>=0)tables.push({rows:rows.slice(hi+1),headers:rows[hi],size:block.length});}
tables.sort((a,b)=>a.size-b.size);let count=0; for(const t of tables)for(const row of t.rows){if(!row.some(v=>/差旅费/.test(v)))continue;let obj={明细类型:t.headers.includes('结算金额')?'携程公司垫付':'个人报销',流程ID:rec.id,流程标题:rec.title,来源链接:rec.url};t.headers.forEach((h,i)=>{if(h)obj[h]=row[i]??''});output.push(obj);count++;}checks.push({...rec,count});}
await fs.writeFile('expense_rows.json',JSON.stringify(output,null,2));await fs.writeFile('expense_checks.json',JSON.stringify(checks,null,2));console.log(JSON.stringify({rows:output.length,empty:checks.filter(x=>!x.count).map(x=>({id:x.id,title:x.title}))}));



