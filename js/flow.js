function Flow(root,cfg){
 const st={files:[],busy:false,cancel:false};root._st=st;
 root.innerHTML=`<div class="prog"><i></i></div><ol class="steps"><li>Select</li><li>Upload</li><li>Settings</li><li>Download</li></ol>
 <div class="g2"><div><div class="drop" tabindex="0" role="button" aria-label="Upload files"><div class="dzi">${ic(IC.upl)}</div><strong>Drag &amp; drop files here or <u>Select files</u></strong><div class="sm">Accepted: ${cfg.accept.join(', ')}${cfg.paste?' · or paste text below':''}</div><input type="file" hidden ${cfg.multi?'multiple':''} accept="${cfg.accept.join(',')}"></div>
 ${cfg.paste?'<div class="row" style="margin-top:10px"><label>Or paste text</label><textarea class="pt"></textarea></div>':''}
 <div class="acts"></div><ul class="list" aria-live="polite"></ul><p class="sm sum"></p></div>
 <div class="card set"><h3 style="margin-bottom:12px">Settings</h3><div class="row"><label>Output file name</label><input type="text" name="out" value="${cfg.out||'merged'}-${stamp()}"></div>${cfg.settings||''}<p class="sm est"></p><button class="btn s pv" style="width:100%;margin-bottom:8px">${ic(IC.eye)}Preview</button><button class="btn go" style="width:100%">${cfg.btn||'Convert now'}</button></div></div><div class="res" aria-live="polite"></div>`;
 const inp=$('input[type=file]',root),drop=$('.drop',root),list=$('.list',root);
 const step=n=>{$$('.steps li',root).forEach((l,i)=>{l.className=i+1<n?'done':i+1===n?'on':'';l.textContent=['Select','Upload','Settings','Download'][i]});$('.prog i',root).style.width=n*25+'%';if(window.innerWidth<640)$$('.steps li',root)[n-1].textContent='Step '+n+' of 4'};
 const opts=()=>{const o={};$$('.set [name]',root).forEach(e=>o[e.name]=e.type==='checkbox'?e.checked:e.type==='radio'?(e.checked?e.value:o[e.name]):e.value);return o};
 async function add(fl){
  for(const f of fl){
   const ext='.'+f.name.split('.').pop().toLowerCase();const it={f,name:f.name,size:f.size,pages:'',ok:true,msg:'',range:'',rot:'0'};
   if(!cfg.accept.includes(ext)){it.ok=false;it.msg='Wrong file type. Use '+cfg.accept.join(', ')}
   else if(!f.size){it.ok=false;it.msg='This file is empty.'}
   else if(f.size>MAX){it.ok=false;it.msg='File is too large. The limit is '+fmt(MAX)+'.'}
   else if(/^\.(docx|pptx|xlsx)$/.test(ext)&&!(b=>b[0]===0x50&&b[1]===0x4b)(new Uint8Array(await f.slice(0,4).arrayBuffer()))){it.ok=false;it.msg='This is not a valid '+ext+' file. It may be an old format or damaged. Open it and save it as '+ext+' again.'}
   else if(ext==='.pdf'){try{const d=await PDFDocument.load(await rd(f));it.pages=d.getPageCount()+' pages';it.n=d.getPageCount()}catch(e){it.ok=false;it.msg=/encrypt/i.test(e.message)?'This PDF is password protected. Enter the password to continue (password support needs the server engine).':'This PDF looks damaged and could not be read.'}}
   if(!cfg.multi)st.files=[];st.files.push(it);if(!it.ok)toast(it.msg)}
  draw()}
 function draw(){
  list.replaceChildren(...st.files.map((it,i)=>{
   const k=h('li',{class:it.ok?'':'err',draggable:cfg.multi?'true':'false'});
   k.append(h('span',{class:'sm','aria-hidden':'true',html:ic(IC.grip)}),h('span',{html:(x=>ftile(x.toUpperCase().replace('JPEG','JPG'),COL[cfg.tag]||COL[{jpg:'img',jpeg:'img',png:'img'}[x]||x]||'#64748b'))(it.name.split('.').pop().toLowerCase())}),h('span',{class:'nm',title:it.name},it.name),h('span',{class:'sm'},fmt(it.size)+(it.pages?' · '+it.pages:'')),h('span',{class:it.ok?'ok':'bad'},it.ok?'ready':'error'));
   if(it.ok&&cfg.perFile&&it.n){const r=h('input',{type:'text',placeholder:'Pages: all or 1-3,5','aria-label':'Pages for '+it.name,style:'width:140px;min-height:36px;border:1px solid var(--bd);border-radius:8px;background:var(--bg);padding:0 8px',on:{input:e=>it.range=e.target.value}});r.value=it.range;const s=h('select',{'aria-label':'Rotate',style:'min-height:36px;border:1px solid var(--bd);border-radius:8px;background:var(--bg)',on:{change:e=>it.rot=e.target.value}},...['0','90','180','270'].map(v=>h('option',{value:v},v==='0'?'No rotation':v+'°')));s.value=it.rot;k.append(r,s)}
   if(!it.ok)k.append(h('div',{class:'bad sm',style:'flex-basis:100%'},it.msg));
   if(cfg.multi){k.append(h('button',{class:'ic','aria-label':'Move up',html:ic(IC.up),on:{click:()=>mv(i,-1)}}),h('button',{class:'ic','aria-label':'Move down',html:ic(IC.down),on:{click:()=>mv(i,1)}}));
    k.ondragstart=e=>e.dataTransfer.setData('i',i);k.ondragover=e=>e.preventDefault();k.ondrop=e=>{e.preventDefault();const a=+e.dataTransfer.getData('i');st.files.splice(i,0,st.files.splice(a,1)[0]);draw()}}
   k.append(h('button',{class:'ic','aria-label':'Remove '+it.name,html:ic(IC.x),on:{click:()=>{st.files.splice(i,1);draw();toast('File removed')}}}));return k}));
  const ok=st.files.filter(x=>x.ok);$('.sum',root).textContent=st.files.length?`${ok.length} file${ok.length>1?'s':''} · ${fmt(ok.reduce((a,x)=>a+x.size,0))}${ok.some(x=>x.n)?' · '+ok.reduce((a,x)=>a+(x.n||0),0)+' pages':''}`:'';
  $('.est',root).textContent=ok.length?'Estimated size: about '+fmt(ok.reduce((a,x)=>a+x.size,0)):'';
  $('.acts',root).replaceChildren(...(cfg.multi&&st.files.length?[['Sort A-Z',()=>st.files.sort((a,b)=>a.name.localeCompare(b.name))],['Sort Z-A',()=>st.files.sort((a,b)=>b.name.localeCompare(a.name))],['Reverse order',()=>st.files.reverse()],['Add more files',()=>inp.click(),1]].map(([t,f,n])=>h('button',{class:'btn s',on:{click:()=>{f();if(!n)draw()}}},t)):[]),...(st.files.length?[h('button',{class:'btn g',on:{click:async()=>{if(await ask('Clear all files?','This removes every file from the list.')){st.files=[];draw()}}}},'Clear all')]:[]));
  const show=ok.length||(cfg.paste&&$('.pt',root)?.value.trim());$('.set',root).style.display=show?'':'none';$('.g2',root).classList.toggle('one',!show);step(show?3:2);cfg.onFiles&&cfg.onFiles(root,st)}
 const mv=(i,d)=>{const j=i+d;if(j<0||j>=st.files.length)return;[st.files[i],st.files[j]]=[st.files[j],st.files[i]];draw()};
 drop.onclick=()=>inp.click();drop.onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();inp.click()}};
 inp.onchange=()=>{add([...inp.files]);inp.value=''};
 drop.ondragover=e=>{e.preventDefault();drop.classList.add('over')};drop.ondragleave=()=>drop.classList.remove('over');drop.ondrop=e=>{e.preventDefault();drop.classList.remove('over');add([...e.dataTransfer.files])};
 const res=$('.res',root);
 const go=async(prev)=>{
  const ok=st.files.filter(x=>x.ok);const pv0=cfg.paste&&$('.pt',root)?.value.trim();if(!ok.length&&pv0)ok.push({f:new File([pv0],'pasted'+cfg.pasteExt),name:'pasted'});if(!ok.length&&!(cfg.paste&&$('.pt',root)?.value.trim())){res.innerHTML='<div class="card"><strong class="bad">Add a file first.</strong><div class="sm">Select or drop at least one valid file to continue.</div></div>';return}
  if(cfg.multi&&cfg.min&&ok.length<cfg.min){res.innerHTML=`<div class="card"><strong class="bad">Add at least ${cfg.min} files.</strong></div>`;return}
  st.cancel=false;step(4);res.innerHTML='<div class="card"><strong class="step">Reading files...</strong><div class="bar" role="progressbar"><i></i></div><button class="btn s cx">Cancel</button></div>';
  $('.cx',res).onclick=()=>{st.cancel=true};const pr=(p,t)=>{$('.bar i',res).style.width=p+'%';$('.step',res).textContent=t;if(st.cancel)throw Error('Cancelled.')};
  try{await new Promise(r=>setTimeout(r,50));let out=await cfg.run(ok,opts(),pr,root);let pvEl=null;if(prev){try{pvEl=await previewEl(out)}catch(e){pvEl=h('p',{class:'bad sm'},'The preview could not be drawn. Your file is fine; download it to open it.')}}
   const o=opts();if(out.length>1&&o.zip!==false&&o.zip!==undefined||out.length>1&&cfg.zipAlways){const z=new JSZip();out.forEach(x=>z.file(x.name,x.blob));out=[{name:(o.out||'files')+'.zip',blob:await z.generateAsync({type:'blob'})}]}
   const c=h('div',{class:'card'},h('strong',{class:'ok dn',html:ic(IC.check)+' Done'}));
   out.forEach(x=>{const a=dl(x.blob,x.name);a.className='btn';a.innerHTML=ic(IC.dl)+'Download '+esc(x.name)+' ('+fmt(x.blob.size)+')';a.style.margin='10px 8px 0 0';c.append(h('div',{},a));if(x.note)c.append(h('p',{class:'sm'},x.note));if(x.stats)c.append(cmpEl(x.stats))});
   const pvBtn=pvEl?null:h('button',{class:'btn s',html:ic(IC.eye)+'Preview',on:{click:async e=>{const b=e.currentTarget;b.disabled=true;b.textContent='Loading preview...';const p=await previewEl(out).catch(()=>h('p',{class:'bad sm'},'The preview could not be drawn.'));b.remove();c.append(p)}}});
   c.append(h('div',{class:'acts'},...(pvBtn?[pvBtn]:[]),h('button',{class:'btn s',on:{click:()=>{res.innerHTML='';step(3)}}},cfg.btn==='Merge now'?'Merge more files':'Convert more'),h('button',{class:'btn g',on:{click:async()=>{if(await ask('Start over?','Your files will be removed from the list.')){st.files=[];draw();res.innerHTML=''}}}},'Start over')));if(pvEl)c.append(pvEl);res.replaceChildren(c);toast('Done')
  }catch(e){res.innerHTML=`<div class="card" style="border-color:var(--er)"><strong class="bad">${esc(e.message)}</strong><div class="acts"><button class="btn s">Try again</button><a class="btn g" href="#footer">Help</a></div></div>`;$('.btn.s',res).onclick=()=>$('.go',root).click();step(3)}};$('.go',root).onclick=()=>go(false);$('.pv',root).onclick=()=>go(true);
 $('.pt',root)?.addEventListener('input',draw);draw();step(2);return st}

