// Convert from PDF tools that run in the browser (pdf.js, Tesseract.js OCR, JSZip, PptxGenJS)
const loadLib=async(g,urls)=>{if(window[g])return;for(const u of urls){try{await new Promise((ok,no)=>{const e=document.createElement('script');e.src=u;e.onload=ok;e.onerror=no;document.head.append(e)});if(window[g])return}catch(e){}}throw Error('Could not load a helper library ('+g+'). Check your internet connection and try again.')};
const XLSX_URLS=['https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js','https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'];
const loadPdf=async f=>{await loadLib('pdfjsLib',['https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js','https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js']);pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';return pdfjsLib.getDocument({data:new Uint8Array(await rd(f.f))}).promise};
const rt=r=>r.cells.map(c=>c.t).join(' ');
function rowsOf(items){const rows=[];items.sort((a,b)=>b.y-a.y||a.x-b.x);
 for(const it of items){let r=rows.find(r=>Math.abs(r.y-it.y)<=Math.max(2,it.size*.45));if(!r)rows.push(r={y:it.y,size:it.size,parts:[]});r.parts.push(it);r.size=Math.max(r.size,it.size)}
 return rows.sort((a,b)=>b.y-a.y).map(r=>{r.parts.sort((a,b)=>a.x-b.x);const cells=[];let cur=null,pv=null;
  for(const q of r.parts){const gap=pv?q.x-(pv.x+pv.w):0;if(pv&&gap>r.size*2.2)cur=null;if(!cur)cells.push(cur={x:q.x,t:''});else if(gap>1.5&&!/\s$/.test(cur.t)&&!/^\s/.test(q.s))cur.t+=' ';cur.t+=q.s;pv=q}
  cells.forEach(c=>c.t=c.t.replace(/\s+/g,' ').trim());return{y:r.y,size:r.size,cells:cells.filter(c=>c.t)}}).filter(r=>r.cells.length)}
async function pageRows(pg){const c=await pg.getTextContent();const items=c.items.filter(i=>i.str.trim()).map(i=>({x:i.transform[4],y:i.transform[5],w:i.width,s:i.str,size:Math.round(Math.abs(i.transform[3])||Math.abs(i.transform[0])||10)}));
 let rows=rowsOf(items.slice());const two=rows.filter(r=>r.cells.length===2&&r.cells.every(c=>c.t.length>25));
 if(rows.length>8&&two.length>rows.length*.35){const xs=two.map(r=>r.cells[1].x).sort((a,b)=>a-b),b=xs[xs.length>>1]-4;rows=[...rowsOf(items.filter(i=>i.x<b)),...rowsOf(items.filter(i=>i.x>=b))]}
 return rows}
let ocr=null,ocrLog=null;
async function ocrRows(pg,lang){await loadLib('Tesseract',['https://cdn.jsdelivr.net/npm/tesseract.js@5.1.0/dist/tesseract.min.js']);if(!ocr||ocr.lang!==lang){if(ocr)await ocr.w.terminate();ocr={lang,w:await Tesseract.createWorker(lang,1,{logger:m=>ocrLog&&ocrLog(m)})}}
 const{data}=await ocr.w.recognize(await renderPage(pg,2.5));const ls=data.lines&&data.lines.length?data.lines:data.text.split('\n').map((t,i)=>({text:t,bbox:{x0:0,y0:i*20,y1:i*20+12}}));
 return ls.map(l=>({ocr:1,y:-l.bbox.y0,size:Math.max(1,l.bbox.y1-l.bbox.y0),cells:[{x:l.bbox.x0,t:l.text.replace(/\s+/g,' ').trim()}]})).filter(r=>r.cells[0].t)}
const listOf=t=>{let m;if(m=t.match(/^[•▪●◦○■·–]\s*(.+)/))return{k:'ul',t:m[1]};if(m=t.match(/^(\d{1,3}|[a-zA-Z])[.)]\s+(.+)/))return{k:'ol',t:m[2],n:m[1]};return null};
function toBlocks(pages){const n=pages.length;
 if(n>2){const cnt={},key=r=>rt(r).replace(/\d+/g,'#');pages.forEach(pg=>new Set([pg[0],pg[pg.length-1]].filter(Boolean).map(key)).forEach(k=>cnt[k]=(cnt[k]||0)+1));pages=pages.map(pg=>pg.filter((r,i)=>(i>0&&i<pg.length-1)||cnt[key(r)]<n*.5))}
 const sz={};pages.flat().forEach(r=>sz[r.size]=(sz[r.size]||0)+rt(r).length);const body=+(Object.entries(sz).sort((a,b)=>b[1]-a[1])[0]||[12])[0],bl=[];let tb=null;
 pages.forEach(pg=>{let pv=null;tb=null;pg.forEach(r=>{const t=rt(r);
  if(r.cells.length>=2&&r.cells.every(c=>c.t.length<=60)){if(tb&&tb.cols===r.cells.length)tb.rows.push(r.cells.map(c=>c.t));else{tb={k:'table',cols:r.cells.length,rows:[r.cells.map(c=>c.t)]};bl.push(tb)}pv=null;return}
  tb=null;const f=r.ocr?1.3:1,lvl=r.size>body*1.8*f?1:r.size>body*1.45*f?2:r.size>body*1.2*f?3:0,li=lvl?null:listOf(t);
  if(pv&&!lvl&&!pv.lvl&&!li&&pv.k==='p'&&pv.y-r.y<=r.size*1.7){const b=bl[bl.length-1];b.t=/[a-z]-$/.test(b.t)&&/^[a-z]/.test(t)?b.t.slice(0,-1)+t:b.t+' '+t}
  else bl.push(lvl?{k:'h',lvl,t}:li?{k:li.k,t:li.t,n:li.n}:{k:'p',t});pv={y:r.y,lvl,k:bl[bl.length-1].k}})});
 if(!bl.length)throw Error('No text found in this PDF. Set "Scanned pages (OCR)" to Auto or Always to read scanned pages.');
 return bl.map(b=>b.k==='table'&&b.rows.length<2?{k:'p',t:b.rows[0].join('  ')}:b)}
const mdText=bl=>{const o=[],c=s=>s.replace(/\|/g,'\\|');bl.forEach(b=>{if(b.k==='h')o.push('\n'+'#'.repeat(b.lvl)+' '+b.t+'\n');else if(b.k==='ul')o.push('- '+b.t);else if(b.k==='ol')o.push((/^\d+$/.test(b.n)?b.n:'1')+'. '+b.t);
 else if(b.k==='table'){const R=b.rows;o.push('\n| '+R[0].map(c).join(' | ')+' |\n|'+' --- |'.repeat(R[0].length)+'\n'+R.slice(1).map(r=>'| '+r.map(c).join(' | ')+' |').join('\n')+'\n')}else o.push('\n'+b.t+'\n')});return o.join('\n').replace(/\n{3,}/g,'\n\n').trim()+'\n'};
async function readPdf(f,o,pr){const doc=await loadPdf(f),n=doc.numPages,mode=o.ocr||'auto',pages=[];
 try{for(let p=1;p<=n;p++){const pg=await doc.getPage(p);let rows=mode==='always'?[]:await pageRows(pg);const chars=rows.reduce((a,r)=>a+rt(r).length,0);
  if(mode==='always'||(mode==='auto'&&chars<20)){const msg='Scanning page '+p+' of '+n+' with OCR (first run downloads language data)...';pr(10+70*(p-1)/n,msg);ocrLog=m=>{if(m.status==='recognizing text')pr(10+70*(p-1+m.progress)/n,msg)};rows=await ocrRows(pg,o.lang||'eng')}
  else pr(10+70*p/n,'Reading page '+p+' of '+n+'...');pages.push(rows)}}
 finally{if(ocr){ocr.w.terminate();ocr=null}}
 pr(85,'Organizing text...');return toBlocks(pages)}
async function makeDocx(bl){const e=s=>s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c])),S={1:40,2:32,3:26};
 const para=(t,b,sz,ind)=>'<w:p>'+(ind?'<w:pPr><w:ind w:left="360"/></w:pPr>':'')+'<w:r>'+(b?'<w:rPr><w:b/><w:sz w:val="'+sz+'"/></w:rPr>':'')+'<w:t xml:space="preserve">'+e(t)+'</w:t></w:r></w:p>';
 const tbl=b=>'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'+['top','left','bottom','right','insideH','insideV'].map(s=>'<w:'+s+' w:val="single" w:sz="4" w:color="999999"/>').join('')+'</w:tblBorders></w:tblPr><w:tblGrid>'+b.rows[0].map(()=>'<w:gridCol w:w="2000"/>').join('')+'</w:tblGrid>'+b.rows.map(r=>'<w:tr>'+r.map(c=>'<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'+para(c)+'</w:tc>').join('')+'</w:tr>').join('')+'</w:tbl><w:p/>';
 const body=bl.map(b=>b.k==='h'?para(b.t,1,S[b.lvl]):b.k==='ul'?para('• '+b.t,0,0,1):b.k==='ol'?para((b.n||'1')+'. '+b.t,0,0,1):b.k==='table'?tbl(b):para(b.t)).join('');
 const z=new JSZip(),X='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
 z.file('[Content_Types].xml',X+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
 z.file('_rels/.rels',X+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
 z.file('word/document.xml',X+'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'+body+'<w:sectPr/></w:body></w:document>');
 return z.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'})}
async function renderPage(pg,scale){const vp=pg.getViewport({scale}),c=document.createElement('canvas');c.width=vp.width;c.height=vp.height;const x=c.getContext('2d');x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);await pg.render({canvasContext:x,viewport:vp}).promise;return c}
async function pdfToJpg(f,o,pr){const doc=await loadPdf(f),z=new JSZip();for(let p=1;p<=doc.numPages;p++){pr(10+80*p/doc.numPages,'Rendering page '+p+' of '+doc.numPages+'...');const c=await renderPage(await doc.getPage(p),(+o.dpi||150)/72);z.file('page-'+String(p).padStart(2,'0')+'.jpg',await new Promise(r=>c.toBlob(r,'image/jpeg',.9)))}
 pr(95,'Finishing...');return[{name:o.out+'.zip',blob:await z.generateAsync({type:'blob'})}]}
async function pdfToPptx(f,o,pr){await loadLib('PptxGenJS',['https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js','https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundle.js','https://unpkg.com/pptxgenjs@3.12.0/dist/pptxgen.bundle.js']);const doc=await loadPdf(f),pp=new PptxGenJS();let first=true;for(let p=1;p<=doc.numPages;p++){pr(10+80*p/doc.numPages,'Building slide '+p+' of '+doc.numPages+'...');const pg=await doc.getPage(p),v=pg.getViewport({scale:1}),W=v.width/72,H=v.height/72;
 if(first){pp.defineLayout({name:'PDF',width:W,height:H});pp.layout='PDF';first=false}
 const s=pp.addSlide();s.addImage({data:(await renderPage(pg,2)).toDataURL('image/jpeg',.9),x:0,y:0,w:W,h:H});s.addNotes((await pg.getTextContent()).items.map(i=>i.str).join(' '))}
 pr(95,'Finishing...');return[{name:o.out+'.pptx',blob:await pp.write({outputType:'blob'})}]}
const CLIENT={'from-jpg':pdfToJpg,'from-ppt':pdfToPptx,
 'from-word':async(f,o,pr)=>{const b=await readPdf(f,o,pr);pr(92,'Building Word file...');return[{name:o.out+'.docx',blob:await makeDocx(b)}]},
 'from-md':async(f,o,pr)=>[{name:o.out+'.md',blob:new Blob([mdText(await readPdf(f,o,pr))],{type:'text/markdown'})}]};
