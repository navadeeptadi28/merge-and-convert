// Compress JPG/PNG, PDF, PowerPoint and Word in the browser.
// The smallest clear size is measured for each file; the user can then pick any size above it.
const {PDFName,PDFRawStream,PDFArray,decodePDFRawStream}=PDFLib;
const toBytes=(v,u)=>Math.round(parseFloat(String(v).replace(',','.'))*(u==='MB'?1048576:1024));
const loadImg=blob=>new Promise((ok,no)=>{const u=URL.createObjectURL(blob),i=new Image();i.onload=()=>{URL.revokeObjectURL(u);ok(i)};i.onerror=()=>no(Error('An image in this file could not be read.'));i.src=u});
const encode=(img,s,type,q,bg)=>{const c=document.createElement('canvas');c.width=Math.max(1,Math.round(img.naturalWidth*s));c.height=Math.max(1,Math.round(img.naturalHeight*s));const x=c.getContext('2d');if(bg){x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height)}x.drawImage(img,0,0,c.width,c.height);return new Promise(r=>c.toBlob(r,type,q))};
const SC=[1,.9,.8,.7,.6,.5];   // images never go below half of their pixel size, so they stay clear
async function fitJpeg(img,max){for(const s of SC){const top=await encode(img,s,'image/jpeg',.95,true);if(top.size<=max)return{b:top,q:.95,s};
 const low=await encode(img,s,'image/jpeg',.45,true);if(low.size>max)continue;let a=.45,z=.95,best={b:low,q:.45};
 for(let i=0;i<6;i++){const m=(a+z)/2,b=await encode(img,s,'image/jpeg',m,true);if(b.size<=max){best={b,q:m};a=m}else z=m}return{...best,s}}return null}
async function compressImage(file,max,pr){const img=await loadImg(file),png=/png$/i.test(file.name);pr(25,'Finding the best quality...');
 if(png&&max>1)for(const s of SC.slice(0,6)){const b=await encode(img,s,'image/png');if(b.size<=max)return{b,ext:'.png',note:'Lossless PNG'+(s<1?', resized to '+Math.round(s*100)+'% of the original size':'')+'.'}}
 let r=await fitJpeg(img,max),miss=false;if(!r){r={b:await encode(img,.5,'image/jpeg',.45,true),q:.45,s:.5};miss=true}
 return{b:r.b,ext:'.jpg',miss,note:'JPEG quality '+Math.round(r.q*100)+'%'+(r.s<1?', resized to '+Math.round(r.s*100)+'% of the original size':'')+(png?', converted from PNG (transparency becomes white)':'')+'.'}}
/* ---- PDF: shrink the pictures inside, never touch text or vectors ---- */
const LVL=[[1,.8],[1,.7],[.85,.65],[.7,.6],[.55,.55],[.45,.5],[.35,.45],[.3,.4]];
const lvP=t=>[+(1-.7*t).toFixed(3),+(.92-.52*t).toFixed(3)];   // t: 0 = best clarity, 1 = the smallest clear setting
const csN=(ctx,cs)=>{cs=ctx.lookup(cs);if(!cs)return 0;const t=cs.toString();if(t==='/DeviceRGB')return 3;if(t==='/DeviceGray')return 1;
 if(cs instanceof PDFArray&&cs.size()>1&&ctx.lookup(cs.get(0)).toString()==='/ICCBased'){const p=ctx.lookup(cs.get(1)),n=p&&p.dict&&ctx.lookup(p.dict.get(PDFName.of('N'))),v=n&&n.asNumber();return v===1||v===3?v:0}return 0};
async function shrinkOne(ctx,ref,obj,L){const d=obj.dict,N=k=>d.get(PDFName.of(k)),num=k=>{const v=ctx.lookup(N(k));return v&&v.asNumber?v.asNumber():0};
 const w=num('Width'),h=num('Height');if(w*h<40000||N('Decode')||N('Mask'))return;const im=ctx.lookup(N('ImageMask'));if(im&&im.toString()==='true')return;
 const comps=csN(ctx,N('ColorSpace'));if(!comps)return;const fl=ctx.lookup(N('Filter')),names=fl instanceof PDFArray?fl.asArray().map(x=>ctx.lookup(x).toString()):[fl&&fl.toString()],last=names[names.length-1],fn=names.length===1?names[0]:(last==='/DCTDecode'&&names.slice(0,-1).every(n=>/^\/(ASCII85Decode|ASCIIHexDecode|FlateDecode|LZWDecode|RunLengthDecode)$/.test(n))?'dct+':'multi');let src;
 if(fn==='/DCTDecode')src=await createImageBitmap(new Blob([obj.contents],{type:'image/jpeg'}));
 else if(fn==='dct+'){const t=obj.dict.clone(ctx);t.set(PDFName.of('Filter'),ctx.obj(names.slice(0,-1).map(n=>PDFName.of(n.slice(1)))));t.delete(PDFName.of('DecodeParms'));const jpg=decodePDFRawStream(PDFRawStream.of(t,obj.contents)).decode();src=await createImageBitmap(new Blob([jpg],{type:'image/jpeg'}))}
 else if(fn==='/FlateDecode'&&num('BitsPerComponent')===8){const px=decodePDFRawStream(obj).decode();if(px.length<w*h*comps)return;const id=new ImageData(w,h);
  for(let i=0,j=0;i<w*h;i++){if(comps===3){id.data[j++]=px[i*3];id.data[j++]=px[i*3+1];id.data[j++]=px[i*3+2]}else{const g=px[i];id.data[j++]=g;id.data[j++]=g;id.data[j++]=g}id.data[j++]=255}
  src=document.createElement('canvas');src.width=w;src.height=h;src.getContext('2d').putImageData(id,0,0)}else return;
 const nw=Math.max(1,Math.round(w*L[0])),nh=Math.max(1,Math.round(h*L[0])),c=document.createElement('canvas');c.width=nw;c.height=nh;const x=c.getContext('2d');x.fillStyle='#fff';x.fillRect(0,0,nw,nh);x.drawImage(src,0,0,nw,nh);
 const u=new Uint8Array(await(await new Promise(r=>c.toBlob(r,'image/jpeg',L[1]))).arrayBuffer());if(u.length>=obj.contents.length*.95)return;
 const nd=ctx.obj({Type:'XObject',Subtype:'Image',Width:nw,Height:nh,ColorSpace:'DeviceRGB',BitsPerComponent:8,Filter:'DCTDecode'}),sm=N('SMask');if(sm)nd.set(PDFName.of('SMask'),sm);ctx.assign(ref,PDFRawStream.of(nd,u))}
async function shrinkPdf(bytes,L,pr,a,z){const doc=await PDFDocument.load(bytes),ctx=doc.context,items=[];
 for(const[ref,obj]of ctx.enumerateIndirectObjects())if(obj instanceof PDFRawStream&&obj.dict.get(PDFName.of('Subtype'))===PDFName.of('Image'))items.push([ref,obj]);
 for(let i=0;i<items.length;i++){pr(a+(z-a)*i/items.length,'Shrinking pictures ('+(i+1)+' of '+items.length+')...');try{await shrinkOne(ctx,items[i][0],items[i][1],L)}catch(e){}}
 return{u:await doc.save({useObjectStreams:true}),n:items.length}}
async function compressPdf(file,max,pr){const bytes=await rd(file);pr(8,'Reading PDF...');let base;
 try{base=await(await PDFDocument.load(bytes)).save({useObjectStreams:true})}catch(e){throw Error(/encrypt/i.test(e.message)?'This PDF is password protected. Remove the password first.':'This PDF could not be read.')}
 const mk=(u,note,miss)=>({b:new Blob([u],{type:'application/pdf'}),note,miss}),ds=L=>'Pictures re-saved at '+Math.round(L[1]*100)+'% quality and '+Math.round(L[0]*100)+'% size. Text and vector graphics are untouched.';
 if(base.length<=max)return mk(base,'Structure optimized. Text stays selectable.');
 pr(12,'Testing the smallest clear setting...');const hi0=LVL.length-1,w=await shrinkPdf(bytes,LVL[hi0],pr,12,40);
 if(w.u.length>=base.length*.98)return mk(base,'This PDF has few large pictures, so it cannot be reduced much.',true);
 if(w.u.length>max)return mk(w.u,ds(LVL[hi0]),true);
 let lo=0,hi=1,pick={u:w.u,L:LVL[hi0]};for(let i=0;i<8;i++){const mid=(lo+hi)/2,L=lvP(mid),s=await shrinkPdf(bytes,L,pr,40+50*i/8,40+50*(i+1)/8);if(s.u.length<=max){pick={u:s.u,L};hi=mid}else lo=mid}
 return mk(pick.u,ds(pick.L))}
/* ---- PowerPoint / Word: re-encode the pictures in the media folder ---- */
async function compressOffice(file,max,pr){const z0=await JSZip.loadAsync(await rd(file)),keys=Object.keys(z0.files),raw={},imgs={};
 const mime=/docx$/i.test(file.name)?'application/vnd.openxmlformats-officedocument.wordprocessingml.document':/xlsx$/i.test(file.name)?'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':'application/vnd.openxmlformats-officedocument.presentationml.presentation';
 for(const n of keys){const e=z0.files[n];if(e.dir)continue;const u=await e.async('uint8array');if(/\/media\/[^/]+\.(jpe?g|png)$/i.test(n)&&u.length>=20000){const png=/png$/i.test(n);imgs[n]={u,png,img:await loadImg(new Blob([u],{type:png?'image/png':'image/jpeg'}))}}else raw[n]=u}
 const build=async L=>{const z=new JSZip();for(const n of keys){if(z0.files[n].dir)continue;let u=raw[n];if(!u){const m=imgs[n];u=m.u;if(L){const nb=new Uint8Array(await(await encode(m.img,L[0],m.png?'image/png':'image/jpeg',L[1],false)).arrayBuffer());if(nb.length<u.length)u=nb}}z.file(n,u)}
  return z.generateAsync({type:'blob',mimeType:mime,compression:'DEFLATE',compressionOptions:{level:9}})};
 pr(20,'Re-packing file...');const base=await build(0);if(base.size<=max)return{b:base,note:'Pictures were not touched; the file was re-packed.'};
 if(!Object.keys(imgs).length)return{b:base,miss:true,note:'This file has no large pictures to shrink.'};
 const LV=[[1,.9],[1,.8],[1,.7],[.85,.7],[.75,.65],[.6,.6],[.5,.55],[.4,.5],[.3,.45]],t=L=>'Pictures re-saved at '+Math.round(L[1]*100)+'% quality and '+Math.round(L[0]*100)+'% size. Text is untouched.';
 let lo=0,hi=1;pr(40,'Testing the smallest clear setting...');let pick={b:await build(LV[LV.length-1]),L:LV[LV.length-1]};if(pick.b.size>max)return{b:pick.b,miss:true,note:t(pick.L)};
 for(let i=0;i<8;i++){const mid=(lo+hi)/2,L=[+(1-.7*mid).toFixed(3),+(.9-.45*mid).toFixed(3)];pr(45+45*i/8,'Finding the best quality...');const b=await build(L);if(b.size<=max){pick={b,L};hi=mid}else lo=mid}return{b:pick.b,note:t(pick.L)}}
const compressAny=(file,max,pr)=>{const e=file.name.split('.').pop().toLowerCase();return /^(jpe?g|png|webp)$/.test(e)?compressImage(file,max,pr):e==='pdf'?compressPdf(file,max,pr):compressOffice(file,max,pr)};
/* ---- smallest possible size, measured once per file ---- */
const AN=new WeakMap();
function analyze(file,pr=()=>{}){if(!AN.has(file))AN.set(file,compressAny(file,1,pr).then(r=>({orig:file.size,floor:Math.min(r.b.size,file.size),r})).catch(e=>{AN.delete(file);throw e}));return AN.get(file)}
function showAn(root,a){const ci=$('.ci',root),pre=$('.pre',root),{orig,floor}=a;
 if(floor>=orig*.97){ci.textContent='Original '+fmt(orig)+'. This file cannot be made noticeably smaller without losing content (smallest '+fmt(floor)+').';pre.replaceChildren();$('.rw',root).style.display='none';return}
 ci.textContent='Original '+fmt(orig)+'. Smallest possible while it stays clear: '+fmt(floor)+'. Pick any size from '+fmt(floor)+' up.';
 const K=1024,M=1048576,rng=$('.rng',root),inp=$('[name=size]',root),un=$('[name=unit]',root),cl=$('.clar',root),CL=[[.6,'Excellent','ok'],[.3,'High','ok'],[.15,'Good','ok'],[.07,'Fair','warn'],[0,'Low','bad']];
 const clar=v=>{const r=v/orig,c=CL.find(x=>r>=x[0]);cl.innerHTML='<span class="'+c[2]+'">&#9679;</span> Estimated clarity: <b>'+c[1]+'</b>'};
 const put=v=>{const mb=v>=M;un.value=mb?'MB':'KB';inp.value=mb?Math.ceil(v/M*100)/100:Math.ceil(v/K);rng.value=v;clar(v)};
 $('.rw',root).style.display='';rng.min=floor;rng.max=Math.max(floor+K,Math.round(orig*.95));rng.step=K;rng.value=Math.round((floor+orig)/2);clar(+rng.value);rng.oninput=()=>put(+rng.value);
 const sync=()=>{const v=toBytes(inp.value,un.value);if(v>0){const c=Math.min(Math.max(v,floor),orig);rng.value=c;clar(c)}};inp.oninput=sync;un.onchange=sync;
 const c=new Map([[floor,'Smallest'],[Math.round(orig/4),'Quarter'],[Math.round(orig/2),'Half']]);[100*K,200*K,500*K,M,2*M,5*M].forEach(v=>c.has(v)||c.set(v,''));
 const L=[...c].filter(([v])=>v>=floor&&v<orig*.95).sort((a,b)=>a[0]-b[0]).filter(([v],i,r)=>!i||v>r[i-1][0]*1.15).slice(0,6);
 pre.replaceChildren(...L.map(([v,l])=>h('button',{type:'button',on:{click:()=>put(v)}},(l?l+' · ':'')+fmt(v))))}
async function compressRun(fs,o,pr){const f=fs[0],file=f.f,orig=file.size,ext=file.name.split('.').pop().toLowerCase(),keep=n=>[{name:o.out+'.'+ext,blob:file,note:n}];
 const target=toBytes(o.size,o.unit);if(!(target>0))throw Error('Type the size you want, for example 500 and KB.');
 if(target>=orig)return keep('Your size ('+fmt(target)+') is not smaller than the file ('+fmt(orig)+'), so it was left unchanged.');
 pr(5,'Checking the smallest possible size...');const an=await analyze(file,pr);
 if(an.floor>=orig*.97)return keep('This file cannot be made noticeably smaller without losing content (smallest '+fmt(an.floor)+'), so it was left unchanged.');
 let msg='',r;if(target<=an.floor){r=an.r;msg='Your size is below the smallest possible, so '+fmt(an.floor)+' was used. '}else r=await compressAny(file,target,pr);
 if(r.b.size>=orig)return keep('This file cannot be made smaller without losing content, so the original was kept.');
 return[{name:o.out+(r.ext||'.'+ext),blob:r.b,stats:{orig,size:r.b.size},note:msg+fmt(orig)+' to '+fmt(r.b.size)+' ('+Math.round(100-r.b.size/orig*100)+'% smaller). '+r.note}]}
const CSET='<div class="note ci" aria-live="polite">Choose a file to see the smallest possible size.</div><div class="row"><label>Size you want</label><div class="two"><input type="text" inputmode="decimal" name="size" placeholder="e.g. 500"><select name="unit" aria-label="Unit"><option>KB</option><option>MB</option></select></div></div><div class="rw" style="display:none"><input type="range" class="rng" aria-label="Size slider"><div class="clar" aria-live="polite"></div></div><div class="chips pre" style="margin-bottom:14px"></div>';
function buildCompress(){Flow($('#cflow'),{accept:['.jpg','.jpeg','.png','.webp','.pdf','.pptx','.docx','.xlsx'],multi:false,btn:'Compress now',out:'compressed',settings:CSET,run:compressRun,
 onFiles:(root,st)=>{const it=st.files.find(x=>x.ok),f=it&&it.f;if(f===root._f)return;root._f=f;const ci=$('.ci',root);$('.pre',root).replaceChildren();$('.rw',root).style.display='none';
  if(!f){ci.textContent='Choose a file to see the smallest possible size.';return}ci.textContent='Checking the smallest possible size...';
  analyze(f,(p,t)=>{if(root._f===f&&t)ci.textContent=t}).then(a=>{if(root._f===f)showAn(root,a)}).catch(e=>{if(root._f===f)ci.textContent=e.message})}})}
const cmpEl=s=>h('div',{class:'cmp'},h('div',{class:'cr'},h('span',{},'Original'),h('div',{class:'cb'},h('i',{style:'width:100%'})),h('b',{},fmt(s.orig))),h('div',{class:'cr'},h('span',{},'Compressed'),h('div',{class:'cb'},h('i',{class:'new',style:'width:'+Math.max(2,s.size/s.orig*100)+'%'})),h('b',{},fmt(s.size))),h('div',{class:'ok',style:'font-weight:600'},'Saved '+Math.round(100-s.size/s.orig*100)+'%'));
