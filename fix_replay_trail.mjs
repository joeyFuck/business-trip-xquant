import fs from 'node:fs/promises';
const file='outputs/travel_visual/template.html';let html=await fs.readFile(file,'utf8');
const replace=(from,to)=>{if(!html.includes(from))throw Error('Missing target');html=html.replace(from,to);};
replace("const path=animationLayer.append('path').attr('class','replay-path').attr('d',routePath(leg.from,leg.to));const length=path.node().getTotalLength();path.attr('stroke-dasharray',length+' '+length).attr('stroke-dashoffset',length);", "const measure=animationLayer.append('path').attr('d',routePath(leg.from,leg.to)).attr('fill','none').attr('stroke','none');const length=measure.node().getTotalLength();const path=animationLayer.append('path').attr('class','replay-path').attr('stroke-linecap','round').attr('stroke-linejoin','round');");
replace('animation.nodes.push({path,icon,length,leg,mode});','animation.nodes.push({path,measure,icon,length,leg,mode});');
replace("node.path.attr('stroke-dashoffset',node.length*(1-f)).attr('marker-end',f===1?'url(#selected)':null);", "const travelled=node.length*f;const steps=Math.max(1,Math.ceil(travelled/2));let trail='';if(f>0){for(let j=0;j<=steps;j++){const p=node.measure.node().getPointAtLength(travelled*j/steps);trail+=(j?'L':'M')+p.x+','+p.y;}}node.path.attr('d',trail||null).attr('marker-end',f===1?'url(#selected)':null);");
replace('const point=node.path.node().getPointAtLength(node.length*f);','const point=node.measure.node().getPointAtLength(node.length*f);');
await fs.writeFile(file,html);
