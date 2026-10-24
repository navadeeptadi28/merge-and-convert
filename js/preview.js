// Result preview. Works on phones too: PDF pages are drawn with pdf.js (phone browsers cannot show a PDF inside a page).
// Word uses docx-preview, PowerPoint uses a small slide renderer below, Markdown is shown formatted.
const PV_BOX='border:1px solid var(--bd);border-radius:12px;margin-top:8px;background:var(--sf)';
const DOCX_URLS=['https://cdn.jsdelivr.net/npm/docx-preview@0.3.3/dist/docx-preview.min.js','https://cdnjs.cloudflare.com/ajax/libs/docx-preview/0.3.3/docx-preview.min.js','https://unpkg.com/docx-preview@0.3.3/dist/docx-preview.min.js'];
const MARKED_URLS=['https://cdn.jsdelivr.net/npm/marked@12.0.2/marked.min.js','https://cdnjs.cloudflare.com/ajax/libs/marked/12.0.2/marked.min.js'];
const pvNote=t=>h('p',{class:'sm',style:'margin:6px 0'},t);

/* ---------- PDF ---------- */
async function pdfPreview(blob,max=12){
 const doc=await loadPdf({f:blob}),n=doc.numPages,box=h('div',{style:PV_BOX+';max-height:560px;overflow:auto;padding:8px;background:var(--bg)'}),w=Math.max(260,Math.min(900,(box.clientWidth||340)));
 const out=h('div',{},pvNote(n+' page'+(n>1?'s':'')+(n>max?'. Showing the first '+max+'.':'.')),box);
 (async()=>{await new Promise(r=>setTimeout(r,80));for(let p=1;p<=Math.min(n,max);p++){const pg=await doc.getPage(p),v=pg.getViewport({scale:1}),cw=Math.min(box.clientWidth-18||w,900),s=cw/v.width*Math.min(2,window.devicePixelRatio||1.5),vp=pg.getViewport({scale:s}),c=h('canvas',{style:'display:block;width:100%;max-width:900px;margin:0 auto 10px;border:1px solid var(--bd);background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.08)'});c.width=vp.width;c.height=vp.height;box.append(c);await pg.render({canvasContext:c.getContext('2d'),viewport:vp}).promise}})().catch(e=>box.append(pvNote('Could not draw every page: '+e.message)));
 return out}

/* ---------- Word ---------- */
async function docxPreview(blob){
 await loadLib('docx',DOCX_URLS);
 const box=h('div',{style:PV_BOX+';max-height:560px;overflow:auto;background:#e5e7eb;padding:8px'}),inner=h('div');box.append(inner);
 const out=h('div',{},pvNote('Word preview. Fonts that are not on this device may look different here; the downloaded file is unchanged.'),box);
 setTimeout(async()=>{try{await docx.renderAsync(blob,inner,null,{className:'docx',inWrapper:true,ignoreWidth:false,ignoreHeight:false,breakPages:true,useBase64URL:true});
  const wr=inner.querySelector('.docx-wrapper'),pg=inner.querySelector('section.docx');
  if(wr&&pg){wr.style.padding='0';wr.style.background='transparent';const fit=()=>{const pw=pg.offsetWidth+0,aw=box.clientWidth-16;if(pw>aw&&aw>0){const s=aw/pw;wr.style.transformOrigin='0 0';wr.style.transform='scale('+s+')';inner.style.height=Math.ceil(wr.scrollHeight*s)+'px';inner.style.overflow='hidden'}};fit()}
 }catch(e){inner.replaceChildren(pvNote('This Word file could not be previewed: '+e.message))}},30);
 return out}

/* ---------- Markdown ---------- */
const clean=html=>{const d=new DOMParser().parseFromString(html,'text/html');d.querySelectorAll('script,iframe,object,embed,style,link').forEach(e=>e.remove());d.querySelectorAll('*').forEach(e=>[...e.attributes].forEach(a=>{if(/^on/i.test(a.name)||/^\s*javascript:/i.test(a.value))e.removeAttribute(a.name)}));return d.body.innerHTML};
const miniMd=t=>{const e=s=>s.replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c])),inl=s=>e(s).replace(/`([^`]+)`/g,'<code>$1</code>').replace(/\*\*([^*]+)\*\*/g,'<b>$1</b>').replace(/\*([^*]+)\*/g,'<i>$1</i>').replace(/\[([^\]]+)\]\(([^)]+)\)/g,'<a href="$2">$1</a>');let o='',ul=false,code=false;
 for(const l of t.split(/\r?\n/)){if(/^```/.test(l)){o+=code?'</pre>':'<pre>';code=!code;continue}if(code){o+=e(l)+'\n';continue}let m;const cl=()=>{if(ul){o+='</ul>';ul=false}};
  if(m=l.match(/^(#{1,6})\s+(.*)/)){cl();o+='<h'+m[1].length+'>'+inl(m[2])+'</h'+m[1].length+'>'}else if(m=l.match(/^\s*[-*+]\s+(.*)/)){if(!ul){o+='<ul>';ul=true}o+='<li>'+inl(m[1])+'</li>'}else if(/^---+$/.test(l.trim())){cl();o+='<hr>'}else if(l.trim()){cl();o+='<p>'+inl(l)+'</p>'}else cl()}
 return o+(code?'</pre>':'')+(ul?'</ul>':'')};
async function mdPreview(blob){
 const src=await blob.text();let html;try{await loadLib('marked',MARKED_URLS);html=marked.parse(src)}catch(e){html=miniMd(src)}
 const view=h('div',{class:'mdview',style:PV_BOX+';max-height:520px;overflow:auto;padding:14px 18px;line-height:1.6'});view.innerHTML=clean(html);
 const raw=h('pre',{style:PV_BOX+';max-height:520px;overflow:auto;white-space:pre-wrap;padding:12px;display:none'},src.slice(0,20000));
 const b1=h('button',{class:'btn s','aria-pressed':'true',on:{click:()=>{view.style.display='';raw.style.display='none';b1.setAttribute('aria-pressed','true');b2.setAttribute('aria-pressed','false')}}},'Formatted'),b2=h('button',{class:'btn s','aria-pressed':'false',on:{click:()=>{view.style.display='none';raw.style.display='';b2.setAttribute('aria-pressed','true');b1.setAttribute('aria-pressed','false')}}},'Markdown text');
 return h('div',{},h('div',{class:'acts'},b1,b2),view,raw)}

/* ---------- PowerPoint (approximate slide renderer) ---------- */
async function pptxPreview(blob,max=30){
 const z=await JSZip.loadAsync(blob),A='http://schemas.openxmlformats.org/drawingml/2006/main',P=NS.p,R=NS.r;
 const pres=await zget(z,'ppt/presentation.xml'),prels=await zget(z,'ppt/_rels/presentation.xml.rels');if(!pres||!prels)throw Error('This is not a PowerPoint file.');
 const sz=pres.getElementsByTagNameNS(P,'sldSz')[0],W=+sz.getAttribute('cx'),H=+sz.getAttribute('cy'),tm={};relsOf(prels).forEach(r=>tm[r.getAttribute('Id')]=r.getAttribute('Target'));
 const slides=[...pres.getElementsByTagNameNS(P,'sldId')].map(s=>'ppt/'+tm[s.getAttributeNS(R,'id')].replace(/^\/?(ppt\/)?/,''));
 const themeName=Object.keys(z.files).find(n=>/^ppt\/theme\/theme\d+\.xml$/.test(n)),theme=themeName?await zget(z,themeName):null,TC={};
 if(theme){const cs=theme.getElementsByTagNameNS(A,'clrScheme')[0];if(cs)[...cs.children].forEach(c=>{const k=c.firstElementChild;if(k)TC[c.localName]=k.getAttribute('val')||k.getAttribute('lastClr')})}
 const MAP={tx1:'dk1',bg1:'lt1',tx2:'dk2',bg2:'lt2'};
 const color=el=>{if(!el)return null;const c=[...el.children].find(x=>/^(srgbClr|schemeClr|sysClr)$/.test(x.localName));if(!c)return null;let v=c.localName==='srgbClr'?c.getAttribute('val'):c.localName==='sysClr'?c.getAttribute('lastClr'):TC[MAP[c.getAttribute('val')]||c.getAttribute('val')];if(!v)return null;return '#'+v};
 const kid=(e,n)=>e&&[...e.children].find(c=>c.localName===n),kids=(e,n)=>e?[...e.children].filter(c=>c.localName===n):[];
 const relMap=async p=>{const r=await zget(z,p.replace(/([^/]+)$/,'_rels/$1')+'.rels'),m={};(r?relsOf(r):[]).forEach(x=>m[x.getAttribute('Id')]={t:x.getAttribute('Target'),ty:x.getAttribute('Type')});return m};
 const abs=(from,t)=>{const parts=from.split('/').slice(0,-1);t.split('/').forEach(s=>{if(s==='..')parts.pop();else if(s&&s!=='.')parts.push(s)});return parts.join('/')};
 const urls={};const imgUrl=async p=>{if(!z.file(p))return null;if(!urls[p]){const ext=p.split('.').pop().toLowerCase();if(/^(emf|wmf|tif|tiff)$/.test(ext))return null;urls[p]=URL.createObjectURL(new Blob([await z.file(p).async('uint8array')],{type:{jpg:'image/jpeg',jpeg:'image/jpeg',png:'image/png',gif:'image/gif',svg:'image/svg+xml',webp:'image/webp'}[ext]||'image/png'}))}return urls[p]};
 const xf=sp=>{const x=sp&&(kid(kid(sp,'spPr'),'xfrm')||kid(sp,'xfrm')||kid(kid(sp,'grpSpPr'),'xfrm'));if(!x)return null;const o=kid(x,'off'),e=kid(x,'ext'),co=kid(x,'chOff'),ce=kid(x,'chExt');if(!o||!e)return null;return{x:+o.getAttribute('x'),y:+o.getAttribute('y'),w:+e.getAttribute('cx'),h:+e.getAttribute('cy'),cx:co&&+co.getAttribute('x'),cy:co&&+co.getAttribute('y'),cw:ce&&+ce.getAttribute('cx'),ch:ce&&+ce.getAttribute('cy'),rot:+(x.getAttribute('rot')||0)/60000,fh:x.getAttribute('flipH')==='1',fv:x.getAttribute('flipV')==='1'}};
 const phOf=sp=>{const ph=kid(kid(kid(sp,'nvSpPr')||kid(sp,'nvPicPr'),'nvPr'),'ph');return ph?{t:ph.getAttribute('type')||'body',i:ph.getAttribute('idx')}:null};
 const inherit=(ph,trees)=>{for(const t of trees){const f=[...t.getElementsByTagNameNS(P,'sp')].find(s=>{const q=phOf(s);return q&&(ph.i&&q.i===ph.i||q.t===ph.t||(ph.t==='ctrTitle'&&q.t==='title'))});if(f&&xf(f))return xf(f)}return null};
 const sizeFor=ph=>!ph?1800:/title/i.test(ph.t)?4400:ph.t==='subTitle'?2400:/^(sldNum|dt|ftr)$/.test(ph.t)?1200:2800;
 const EM=12700;
 const rendSlide=async(path)=>{
  const sd=await zget(z,path),rm=await relMap(path),lay=Object.entries(rm).find(([,v])=>/slideLayout$/.test(v.ty)),layP=lay&&abs(path,lay[1].t),layD=layP&&await zget(z,layP),lrm=layP&&await relMap(layP),mst=lrm&&Object.values(lrm).find(v=>/slideMaster$/.test(v.ty)),mstD=mst&&await zget(z,abs(layP,mst.t)),trees=[layD,mstD].filter(Boolean);
  const phStyle=ph=>{if(!ph)return{};for(const t of trees){const f=[...t.getElementsByTagNameNS(P,'sp')].find(s=>{const q=phOf(s);return q&&(ph.i&&q.i===ph.i||q.t===ph.t||(ph.t==='ctrTitle'&&q.t==='title'))});const l=f&&kid(kid(kid(f,'txBody'),'lstStyle'),'lvl1pPr');if(l){const d=kid(l,'defRPr');return{algn:l.getAttribute('algn'),sz:d&&+d.getAttribute('sz')||0}}}
   const ts=mstD&&mstD.getElementsByTagNameNS(P,/title/i.test(ph.t)?'titleStyle':'bodyStyle')[0],l=ts&&kid(ts,'lvl1pPr'),d=l&&kid(l,'defRPr');return l?{algn:l.getAttribute('algn'),sz:d&&+d.getAttribute('sz')||0}:{}};
  const box=h('div',{style:'position:relative;width:100%;aspect-ratio:'+W+'/'+H+';overflow:hidden;background:#fff;container-type:inline-size;border:1px solid var(--bd);border-radius:6px;color:#000;font-family:Calibri,Arial,sans-serif'});
  let bg=null;for(const d of [sd,layD,mstD]){const b=d&&d.getElementsByTagNameNS(P,'bg')[0];if(b){bg=color(kid(kid(b,'bgPr'),'solidFill'));if(bg)break}}if(bg)box.style.background=bg;
  const pct=(v,t)=>(v/t*100).toFixed(3)+'%';
  const place=(el,g,tr)=>{let{x,y,w,h:hh}=g;if(tr){x=tr.ox+(x-tr.cx)*tr.sx;y=tr.oy+(y-tr.cy)*tr.sy;w*=tr.sx;hh*=tr.sy}el.style.cssText+=';position:absolute;box-sizing:border-box;left:'+pct(x,W)+';top:'+pct(y,H)+';width:'+pct(w,W)+';height:'+pct(hh,H);if(g.rot)el.style.transform='rotate('+g.rot+'deg)'+(g.fh?' scaleX(-1)':'')+(g.fv?' scaleY(-1)':'');else if(g.fh||g.fv)el.style.transform=(g.fh?'scaleX(-1) ':'')+(g.fv?'scaleY(-1)':'')};
  const runs=(p,ph,body)=>{const pp=kid(p,'pPr'),st=phStyle(ph),al=pp&&pp.getAttribute('algn')||st.algn||(ph&&/^(ctrTitle|subTitle)$/.test(ph.t)?'ctr':''),d=h('div',{style:'text-align:'+({ctr:'center',r:'right',just:'justify'}[al]||'left')+';white-space:pre-wrap;overflow-wrap:anywhere'});
   const mar=pp&&+pp.getAttribute('marL')||0;if(mar)d.style.marginLeft=pct(mar,W);
   if(pp&&kid(pp,'buChar')){const bc=kid(pp,'buChar').getAttribute('char');d.append(h('span',{},(bc||'•')+' '))}
   const rs=[...p.children].filter(c=>/^(r|fld)$/.test(c.localName));let any=false;
   for(const r of rs){const t=kid(r,'t'),rp=kid(r,'rPr');if(!t||!t.textContent)continue;any=true;const sz=(rp&&+rp.getAttribute('sz'))||st.sz||sizeFor(ph),s=h('span',{},t.textContent);s.style.fontSize=(sz/100*EM/W*100).toFixed(3)+'cqw';if(rp){if(rp.getAttribute('b')==='1')s.style.fontWeight='700';if(rp.getAttribute('i')==='1')s.style.fontStyle='italic';if(/^(sng|dbl)/.test(rp.getAttribute('u')||''))s.style.textDecoration='underline';const c=color(kid(rp,'solidFill'));if(c)s.style.color=c;const lf=kid(rp,'latin');if(lf&&lf.getAttribute('typeface')&&!/^\+/.test(lf.getAttribute('typeface')))s.style.fontFamily='"'+lf.getAttribute('typeface')+'",Calibri,Arial,sans-serif'}d.append(s)}
   if(!any){const s=h('span',{},'\u00a0');s.style.fontSize=((st.sz||sizeFor(ph))/100*EM/W*100).toFixed(3)+'cqw';d.append(s)}return d};
  const shape=async(sp,tr)=>{const ph=phOf(sp);let g=xf(sp)||(ph&&inherit(ph,trees));if(!g)return;const e=h('div',{});place(e,g,tr);const spPr=kid(sp,'spPr'),f=color(kid(spPr,'solidFill')),ln=kid(spPr,'ln'),lc=ln&&color(kid(ln,'solidFill')),geom=kid(spPr,'prstGeom')&&kid(spPr,'prstGeom').getAttribute('prst');
   if(f)e.style.background=f;if(lc)e.style.border=Math.max(1,(+ln.getAttribute('w')||12700)/EM*(box.clientWidth||600)/(W/EM))+'px solid '+lc;if(geom==='ellipse')e.style.borderRadius='50%';else if(geom==='roundRect')e.style.borderRadius='12%';
   const tb=kid(sp,'txBody');if(tb){const bp=kid(tb,'bodyPr'),an=bp&&bp.getAttribute('anchor')||(ph&&/title/i.test(ph.t)?'ctr':'t');e.style.display='flex';e.style.flexDirection='column';e.style.justifyContent={ctr:'center',b:'flex-end'}[an]||'flex-start';const ins=a=>{const v=bp&&bp.getAttribute(a);return v!==null&&v!==undefined?+v:(a==='lIns'||a==='rIns'?91440:45720)};e.style.padding=pct(ins('tIns'),H)+' '+pct(ins('rIns'),W)+' '+pct(ins('bIns'),H)+' '+pct(ins('lIns'),W);e.style.overflow='hidden';kids(tb,'p').forEach(p=>e.append(runs(p,ph)))}
   box.append(e)};
  const pic=async(pc,tr)=>{const g=xf(pc);if(!g)return;const bl=kid(kid(pc,'blipFill'),'blip'),id=bl&&bl.getAttributeNS(R,'embed'),rel=id&&rm[id];if(!rel)return;const u=await imgUrl(abs(path,rel.t));if(!u)return;const e=h('div',{style:'overflow:hidden'}),im=h('img',{src:u,alt:'',style:'position:absolute;display:block'}),sr=kid(kid(pc,'blipFill'),'srcRect'),v=n=>sr?(+sr.getAttribute(n)||0)/100000:0,l=v('l'),t=v('t'),r=v('r'),b=v('b'),fw=1-l-r,fh=1-t-b;place(e,g,tr);im.style.width=(100/fw)+'%';im.style.height=(100/fh)+'%';im.style.left=(-l/fw*100)+'%';im.style.top=(-t/fh*100)+'%';e.append(im);box.append(e)};
  const table=(gf,tr)=>{const g=xf(gf);const tbl=gf.getElementsByTagNameNS(A,'tbl')[0];if(!g||!tbl)return;const e=h('div',{style:'overflow:hidden'});place(e,g,tr);const t=h('table',{style:'width:100%;height:100%;border-collapse:collapse;table-layout:fixed'});kids(tbl,'tr').forEach(r=>{const tr2=h('tr');kids(r,'tc').forEach(c=>{const td=h('td',{style:'border:1px solid #888;padding:2px 4px;vertical-align:top'});const fill=color(kid(kid(c,'tcPr'),'solidFill'));if(fill)td.style.background=fill;kids(kid(c,'txBody'),'p').forEach(p=>td.append(runs(p,null)));tr2.append(td)});t.append(tr2)});e.append(t);box.append(e)};
  const walk=async(tree,tr)=>{for(const c of [...tree.children]){const n=c.localName;if(n==='sp')await shape(c,tr);else if(n==='pic')await pic(c,tr);else if(n==='graphicFrame')table(c,tr);else if(n==='cxnSp'){const g=xf(c),lc=color(kid(kid(kid(c,'spPr'),'ln'),'solidFill'))||'#555';if(g){const e=h('div',{style:'background:'+lc});place(e,g,tr);if(g.w>=g.h)e.style.height='1px';else e.style.width='1px';box.append(e)}}
    else if(n==='grpSp'){const g=xf(c);if(g&&g.cw&&g.ch){const t2={cx:g.cx,cy:g.cy,ox:g.x,oy:g.y,sx:g.w/g.cw,sy:g.h/g.ch};await walk(c,tr?{...t2,ox:tr.ox+(t2.ox-tr.cx)*tr.sx,oy:tr.oy+(t2.oy-tr.cy)*tr.sy,sx:t2.sx*tr.sx,sy:t2.sy*tr.sy}:t2)}else await walk(c)}}};
  const tree=sd.getElementsByTagNameNS(P,'spTree')[0];if(tree)await walk(tree);return box};
 const wrap=h('div',{style:PV_BOX+';max-height:560px;overflow:auto;padding:8px;background:var(--bg)'}),out=h('div',{},pvNote(slides.length+' slide'+(slides.length>1?'s':'')+(slides.length>max?'. Showing the first '+max+'.':'.')+' This is an approximate preview; the downloaded file is unchanged.'),wrap);
 for(let i=0;i<Math.min(slides.length,max);i++){try{const s=await rendSlide(slides[i]);wrap.append(h('div',{style:'margin-bottom:10px'},s,h('div',{class:'sm',style:'text-align:center;margin-top:2px'},'Slide '+(i+1))))}catch(e){wrap.append(pvNote('Slide '+(i+1)+' could not be drawn.'))}}
 return out}

/* ---------- entry point used by the converters ---------- */
async function previewEl(out){
 const x=out[0],w=h('div',{class:'res'});w.append(h('strong',{},'Preview'));
 if(out.length>1)w.append(h('p',{class:'sm'},out.length+' files: '+out.map(o=>o.name).join(', ')+'. Showing the first.'));
 try{
  if(/\.pdf$/i.test(x.name))w.append(await pdfPreview(x.blob));
  else if(/\.docx$/i.test(x.name))w.append(await docxPreview(x.blob));
  else if(/\.pptx$/i.test(x.name))w.append(await pptxPreview(x.blob));
  else if(/\.md$/i.test(x.name))w.append(await mdPreview(x.blob));
  else if(/\.(txt|csv)$/i.test(x.name))w.append(h('pre',{style:PV_BOX+';max-height:320px;overflow:auto;white-space:pre-wrap;padding:12px'},(await x.blob.text()).slice(0,6000)));
  else if(/\.(jpe?g|png|webp)$/i.test(x.name))w.append(h('img',{alt:'Preview',src:URL.createObjectURL(x.blob),style:'max-width:100%;max-height:460px;border:1px solid var(--bd);border-radius:12px;margin-top:8px'}));
  else if(/\.xlsx$/i.test(x.name)){await loadLib('XLSX',XLSX_URLS);const wb=XLSX.read(await x.blob.arrayBuffer()),d=h('div',{style:PV_BOX+';max-height:360px;overflow:auto;padding:8px;font-size:14px'});wb.SheetNames.slice(0,3).forEach(n=>{d.append(h('strong',{},n));const t=h('div');t.innerHTML=clean(XLSX.utils.sheet_to_html(wb.Sheets[n]));d.append(t)});w.append(d)}
  else if(/\.zip$/i.test(x.name)){const z=await JSZip.loadAsync(x.blob),ns=Object.keys(z.files).filter(n=>/\.(jpe?g|png)$/i.test(n)).sort();w.append(pvNote(Object.keys(z.files).length+' files inside. Showing the first '+Math.min(3,ns.length)+' pages.'));for(const n of ns.slice(0,3))w.append(h('img',{alt:n,src:URL.createObjectURL(await z.files[n].async('blob')),style:'max-width:100%;max-height:420px;border:1px solid var(--bd);border-radius:12px;margin:6px 6px 0 0'}))}
  else w.append(pvNote('No visual preview for this file type. Download it to open it.'));
 }catch(e){w.append(h('p',{class:'bad sm'},'The preview could not be drawn ('+e.message+'). Your file is fine; download it to open it.'))}
 return w}
