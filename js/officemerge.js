// Merge Word and PowerPoint files in their own format, in the browser (JSZip + DOM XML).
const NS={w:'http://schemas.openxmlformats.org/wordprocessingml/2006/main',r:'http://schemas.openxmlformats.org/officeDocument/2006/relationships',p:'http://schemas.openxmlformats.org/presentationml/2006/main',pr:'http://schemas.openxmlformats.org/package/2006/relationships'};
const XP=new DOMParser(),XS=new XMLSerializer(),xml=s=>XP.parseFromString(s,'application/xml'),str=d=>XS.serializeToString(d);
const zget=async(z,n)=>z.file(n)?xml(await z.file(n).async('string')):null;
const relsOf=d=>[...d.getElementsByTagNameNS(NS.pr,'Relationship')];
async function openZip(f){try{return await JSZip.loadAsync(await rd(f.f))}catch(e){throw Error(f.name+' is not a valid file. It may be damaged or saved in an old format.')}}
function ctAdd(ct,src){if(!src)return;const have=new Set([...ct.getElementsByTagName('Default')].map(e=>e.getAttribute('Extension').toLowerCase()));
 for(const e of src.getElementsByTagName('Default')){const x=e.getAttribute('Extension').toLowerCase();if(!have.has(x)){ct.documentElement.append(ct.importNode(e,true));have.add(x)}}}
const addRel=(doc,id,type,target,mode)=>{const n=doc.createElementNS(NS.pr,'Relationship');n.setAttribute('Id',id);n.setAttribute('Type',type);n.setAttribute('Target',target);if(mode)n.setAttribute('TargetMode',mode);doc.documentElement.append(n)};
const DOCX='application/vnd.openxmlformats-officedocument.wordprocessingml.document',PPTX='application/vnd.openxmlformats-officedocument.presentationml.presentation';
/* ---------------- Word ---------------- */
async function mergeDocx(fs,o,pr){const zs=[];for(const f of fs)zs.push(await openZip(f));const B=zs[0];
 const bd=await zget(B,'word/document.xml');if(!bd)throw Error(fs[0].name+' is not a Word document.');
 const body=bd.getElementsByTagNameNS(NS.w,'body')[0],lastSect=b=>[...b.children].reverse().find(c=>c.localName==='sectPr'),bsect=lastSect(body),
  brels=await zget(B,'word/_rels/document.xml.rels'),ct=await zget(B,'[Content_Types].xml'),bst=await zget(B,'word/styles.xml'),bnum=await zget(B,'word/numbering.xml');
 const ins=n=>bsect?body.insertBefore(n,bsect.parentNode===body?bsect:null):body.append(n),E=t=>bd.createElementNS(NS.w,'w:'+t);
 const mkp=(...k)=>{const p=E('p');p.append(...k);return p},brk=()=>{const r=E('r'),b=E('br');b.setAttributeNS(NS.w,'w:type','page');r.append(b);return mkp(r)};
 const head=t=>{const r=E('r'),rp=E('rPr'),sz=E('sz'),tx=E('t');sz.setAttributeNS(NS.w,'w:val','32');rp.append(E('b'),sz);tx.textContent=t;r.append(rp,tx);return mkp(r)};
 const sectP=s=>{const pp=E('pPr');pp.append(s);return mkp(pp)},clean=s=>{const c=s.cloneNode(true);[...c.children].filter(e=>/^(header|footer)Reference$/.test(e.localName)).forEach(e=>e.remove());return c};
 const name=f=>f.name.replace(/\.docx$/i,''),keep=o.fmt!=='first';let docPr=1000;
 if(o.heading)body.insertBefore(head(name(fs[0])),body.firstChild);
 if(keep&&bsect&&zs.length>1)ins(sectP(bsect.cloneNode(true)));
 for(let i=1;i<zs.length;i++){pr(10+80*i/zs.length,'Adding '+fs[i].name+'...');const Z=zs[i],sd=await zget(Z,'word/document.xml');if(!sd)throw Error(fs[i].name+' is not a Word document.');
  const sb=sd.getElementsByTagNameNS(NS.w,'body')[0],ssect=lastSect(sb),srels=await zget(Z,'word/_rels/document.xml.rels'),map={},drop=new Set();
  ctAdd(ct,await zget(Z,'[Content_Types].xml'));
  for(const r of srels?relsOf(srels):[]){const id=r.getAttribute('Id'),ty=r.getAttribute('Type'),tg=r.getAttribute('Target');
   if(/\/hyperlink$/.test(ty)){map[id]='mh'+i+'_'+id;addRel(brels,map[id],ty,tg,'External')}
   else if(/\/image$/.test(ty)&&Z.file('word/'+tg)){const nn='media/m'+i+'_'+tg.split('/').pop();B.file('word/'+nn,await Z.file('word/'+tg).async('uint8array'));map[id]='mi'+i+'_'+id;addRel(brels,map[id],ty,nn)}
   else if(!/\/(styles|numbering|settings|fontTable|theme|webSettings|header|footer|footnotes|endnotes|comments|customXml)$/.test(ty))drop.add(id)}
  let offN=0;const snum=await zget(Z,'word/numbering.xml');
  if(snum&&bnum){const bA=[...bnum.getElementsByTagNameNS(NS.w,'abstractNum')],bN=[...bnum.getElementsByTagNameNS(NS.w,'num')],g=(e,a)=>+e.getAttributeNS(NS.w,a),offA=Math.max(-1,...bA.map(a=>g(a,'abstractNumId')))+1;offN=Math.max(0,...bN.map(a=>g(a,'numId')));const first=bN[0];
   for(const a of snum.getElementsByTagNameNS(NS.w,'abstractNum')){const c=bnum.importNode(a,true);c.setAttributeNS(NS.w,'w:abstractNumId',g(a,'abstractNumId')+offA);first?bnum.documentElement.insertBefore(c,first):bnum.documentElement.append(c)}
   for(const n of snum.getElementsByTagNameNS(NS.w,'num')){const c=bnum.importNode(n,true);c.setAttributeNS(NS.w,'w:numId',g(n,'numId')+offN);const ab=c.getElementsByTagNameNS(NS.w,'abstractNumId')[0];if(ab)ab.setAttributeNS(NS.w,'w:val',g(ab,'val')+offA);bnum.documentElement.append(c)}}
  const bumpNum=root=>{if(offN)for(const e of root.getElementsByTagNameNS(NS.w,'numId')){const v=+e.getAttributeNS(NS.w,'val');if(v)e.setAttributeNS(NS.w,'w:val',v+offN)}};
  const sst=await zget(Z,'word/styles.xml');if(sst&&bst){const have=new Set([...bst.getElementsByTagNameNS(NS.w,'style')].map(s=>s.getAttributeNS(NS.w,'styleId')));
   for(const s of sst.getElementsByTagNameNS(NS.w,'style'))if(!have.has(s.getAttributeNS(NS.w,'styleId'))){const c=bst.importNode(s,true);bumpNum(c);bst.documentElement.append(c)}}
  const wrap=bd.createElement('x');[...sb.children].filter(c=>c!==ssect).forEach(c=>wrap.append(bd.importNode(c,true)));bumpNum(wrap);const bad=[];
  for(const el of [...wrap.getElementsByTagName('*')]){for(const a of [...el.attributes])if(a.namespaceURI===NS.r){if(map[a.value])a.value=map[a.value];else if(drop.has(a.value))bad.push(el)}if(el.localName==='docPr')el.setAttribute('id',++docPr)}
  bad.forEach(el=>{let e=el;while(e.parentNode&&e.parentNode!==wrap&&e.localName!=='r')e=e.parentNode;if(e.parentNode)e.parentNode.removeChild(e)});
  if(!keep&&o.pb)ins(brk());if(o.heading)ins(head(name(fs[i])));[...wrap.childNodes].forEach(ins);
  if(keep&&ssect){const last=i===zs.length-1;if(last){const ns=clean(ssect);bd.importNode(ns,true);if(bsect&&bsect.parentNode===body){body.insertBefore(bd.importNode(ns,true),bsect);bsect.remove()}else body.append(bd.importNode(ns,true))}else ins(sectP(bd.importNode(clean(ssect),true)))}}
 B.file('word/document.xml',str(bd));B.file('[Content_Types].xml',str(ct));if(brels)B.file('word/_rels/document.xml.rels',str(brels));if(bst)B.file('word/styles.xml',str(bst));if(bnum)B.file('word/numbering.xml',str(bnum));
 pr(95,'Finishing...');return[{name:o.out+'.docx',blob:await B.generateAsync({type:'blob',mimeType:DOCX,compression:'DEFLATE'})}]}
/* ---------------- PowerPoint (uses the first deck's theme) ---------------- */
async function mergePptx(fs,o,pr){const zs=[];for(const f of fs)zs.push(await openZip(f));const B=zs[0],PN=NS.p;
 const bp=await zget(B,'ppt/presentation.xml'),bpr=await zget(B,'ppt/_rels/presentation.xml.rels'),ct=await zget(B,'[Content_Types].xml');if(!bp||!bpr)throw Error(fs[0].name+' is not a PowerPoint file.');
 const lst=bp.getElementsByTagNameNS(PN,'sldIdLst')[0];if(!lst)throw Error(fs[0].name+' has no slides.');
 const layouts=async z=>{const r=[];for(const n of Object.keys(z.files).filter(n=>/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(n))){const d=xml(await z.file(n).async('string')),c=d.getElementsByTagNameNS(PN,'cSld')[0];r.push({n:n.split('/').pop(),name:c&&c.getAttribute('name'),type:d.documentElement.getAttribute('type')})}return r};
 const order=async z=>{const p=await zget(z,'ppt/presentation.xml'),r=await zget(z,'ppt/_rels/presentation.xml.rels'),m={};if(!p||!r)throw Error('One of the files is not a PowerPoint file.');relsOf(r).forEach(x=>m[x.getAttribute('Id')]=x.getAttribute('Target'));
  return[...p.getElementsByTagNameNS(PN,'sldId')].map(s=>'ppt/'+m[s.getAttributeNS(NS.r,'id')].replace(/^\//,'').replace(/^ppt\//,''))};
 const baseLay=await layouts(B),blank=baseLay.find(l=>/blank/i.test(l.name||''))||baseLay[0],sz=bp.getElementsByTagNameNS(PN,'sldSz')[0],W=+sz.getAttribute('cx'),H=+sz.getAttribute('cy');
 let n=Math.max(0,...Object.keys(B.files).map(x=>+(x.match(/^ppt\/slides\/slide(\d+)\.xml$/)||[0,0])[1])),sid=Math.max(255,...[...lst.children].map(s=>+s.getAttribute('id')));
 const addSlide=(x,rels)=>{n++;B.file('ppt/slides/slide'+n+'.xml',x);
  B.file('ppt/slides/_rels/slide'+n+'.xml.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="'+NS.pr+'">'+rels.map(r=>'<Relationship Id="'+r.Id+'" Type="'+r.Type+'" Target="'+esc(r.Target)+'"'+(r.Mode?' TargetMode="'+r.Mode+'"':'')+'/>').join('')+'</Relationships>');
  const ov=ct.createElementNS(ct.documentElement.namespaceURI,'Override');ov.setAttribute('PartName','/ppt/slides/slide'+n+'.xml');ov.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.presentationml.slide+xml');ct.documentElement.append(ov);
  const rid='rIdMg'+n;addRel(bpr,rid,NS.r+'/slide','slides/slide'+n+'.xml');const sl=bp.createElementNS(PN,'p:sldId');sl.setAttribute('id',++sid);sl.setAttributeNS(NS.r,'r:id',rid);lst.append(sl)};
 const divider=t=>'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="'+NS.r+'" xmlns:p="'+PN+'"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="'+Math.round(W*.1)+'" y="'+Math.round(H*.4)+'"/><a:ext cx="'+Math.round(W*.8)+'" cy="'+Math.round(H*.2)+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="4000" b="1"/><a:t>'+esc(t)+'</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>';
 for(let i=1;i<zs.length;i++){pr(10+80*i/zs.length,'Adding '+fs[i].name+'...');const Z=zs[i],sl=await layouts(Z),ord=await order(Z);ctAdd(ct,await zget(Z,'[Content_Types].xml'));
  if(o.divider)addSlide(divider(fs[i].name.replace(/\.pptx$/i,'')),[{Id:'rId1',Type:NS.r+'/slideLayout',Target:'../slideLayouts/'+blank.n}]);
  for(const path of ord){let x=await Z.file(path).async('string');const rd0=await zget(Z,path.replace('slides/','slides/_rels/')+'.rels'),rels=[],drop=new Set();
   for(const r of rd0?relsOf(rd0):[]){const id=r.getAttribute('Id'),ty=r.getAttribute('Type'),tg=r.getAttribute('Target'),base=tg.split('/').pop();
    if(/\/slideLayout$/.test(ty)){const s=sl.find(l=>l.n===base),m=baseLay.find(l=>s&&l.name&&l.name===s.name)||baseLay.find(l=>s&&l.type&&l.type===s.type)||blank;rels.push({Id:id,Type:ty,Target:'../slideLayouts/'+m.n})}
    else if(/\/image$/.test(ty)&&Z.file('ppt/media/'+base)){const nn='d'+i+'_'+base;B.file('ppt/media/'+nn,await Z.file('ppt/media/'+base).async('uint8array'));rels.push({Id:id,Type:ty,Target:'../media/'+nn})}
    else if(/\/hyperlink$/.test(ty))rels.push({Id:id,Type:ty,Target:tg,Mode:r.getAttribute('TargetMode')||'External'});else drop.add(id)}
   if(drop.size){const d=xml(x);for(const el of [...d.getElementsByTagName('*')])for(const a of [...el.attributes])if(a.namespaceURI===NS.r&&drop.has(a.value)){let e=el;while(e.parentNode&&e.parentNode.nodeType===1&&!/^(pic|graphicFrame|sp)$/.test(e.localName))e=e.parentNode;if(e.parentNode&&e.parentNode.nodeType===1)e.parentNode.removeChild(e)}x=str(d)}
   addSlide(x,rels)}}
 B.file('ppt/presentation.xml',str(bp));B.file('ppt/_rels/presentation.xml.rels',str(bpr));B.file('[Content_Types].xml',str(ct));
 pr(95,'Finishing...');return[{name:o.out+'.pptx',blob:await B.generateAsync({type:'blob',mimeType:PPTX,compression:'DEFLATE'})}]}
