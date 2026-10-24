IC.home='<path fill="currentColor" fill-opacity=".16" stroke="none" d="M4 11 12 4l8 7v8a1.5 1.5 0 0 1-1.5 1.5h-13A1.5 1.5 0 0 1 4 19z"/><path d="M4 11 12 4l8 7v8a1.5 1.5 0 0 1-1.5 1.5h-13A1.5 1.5 0 0 1 4 19zM9.5 20.5V14h5v6.5"/>';
const go2=u=>{location.hash=u};
const ALL=[...MT.map(m=>({n:'Merge '+m[1],a:()=>go2('#merge/'+m[0])})),{n:'Split PDF',a:()=>go2('#split')},{n:'Compress files (image, PDF, PowerPoint, Word)',a:()=>go2('#compress')},...TO.map(t=>({n:t[1],a:()=>go2('#topdf/'+t[0])})),...FROM.map(t=>({n:t[1],a:()=>go2('#frompdf/'+t[0])}))];
const SYN={excel:'xlsx spreadsheet',word:'docx doc',powerpoint:'pptx ppt slides presentation',image:'jpg jpeg png picture photo',compress:'shrink reduce smaller size kb mb',markdown:'md',merge:'combine join',split:'cut extract pages',html:'web page'};
const GRP=[['Merge',IC.merge],['Split',IC.split],['Compress',IC.compress],['PDF to',IC.fromPdf],['to PDF',IC.toPdf]],POP=['Merge PDF','Split PDF','Compress files (image, PDF, PowerPoint, Word)','Word to PDF','PDF to Word'];
ALL.forEach(x=>{const l=x.n.toLowerCase(),g=GRP.find(g=>g[0]==='PDF to'?x.n.startsWith('PDF to'):g[0]==='to PDF'?x.n.endsWith('to PDF'):x.n.startsWith(g[0]))||GRP[0];x.i=g[1];x.g=g[0]==='PDF to'?'Convert from PDF':g[0]==='to PDF'?'Convert to PDF':g[0];x.k=Object.entries(SYN).filter(([w])=>l.includes(w)).map(e=>e[1]).join(' ');x.pop=POP.includes(x.n)});
const q=$('#q'),dd=$('#dd');let sel=0,shown=[];
const mark=i=>{sel=i;$$('button',dd).forEach((b,j)=>b.classList.toggle('on',j===i))},go=x=>{x.a();q.value='';q.blur();dd.style.display='none'};
function renderDD(){const v=q.value.toLowerCase().trim();shown=(v?ALL.filter(x=>(x.n+' '+x.k).toLowerCase().includes(v)):ALL.filter(x=>x.pop)).slice(0,8);sel=0;
 dd.replaceChildren(h('div',{class:'ddh'},v?(shown.length?'Results':'No tools match "'+q.value+'". Try excel, jpg or compress.'):'Popular tools'),...shown.map((x,i)=>h('button',{type:'button',class:i?'':'on',on:{mousedown:e=>{e.preventDefault();go(x)},mousemove:()=>mark(i)}},h('span',{class:'ddi',html:ic(x.i)}),h('span',{class:'ddn'},x.n),h('small',{},x.g))));dd.style.display='block'}
q.addEventListener('focus',renderDD);q.addEventListener('input',renderDD);q.addEventListener('blur',()=>{dd.style.display='none'});
q.addEventListener('keydown',e=>{if(e.key==='ArrowDown'||e.key==='ArrowUp'){e.preventDefault();if(shown.length)mark((sel+(e.key==='ArrowDown'?1:-1)+shown.length)%shown.length)}else if(e.key==='Enter'&&shown[sel]){e.preventDefault();go(shown[sel])}else if(e.key==='Escape')q.blur()});
document.addEventListener('keydown',e=>{if(e.key==='k'&&(e.ctrlKey||e.metaKey)){e.preventDefault();q.focus()}else if(e.key==='/'&&!/INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName)){e.preventDefault();q.focus()}});
const root=document.documentElement;let th='';try{th=localStorage.getItem('mc-theme')||''}catch(e){}if(th)root.dataset.theme=th;
$('#theme').onclick=()=>{const d=root.dataset.theme==='dark'||(!root.dataset.theme&&matchMedia('(prefers-color-scheme:dark)').matches);root.dataset.theme=d?'light':'dark';try{localStorage.setItem('mc-theme',root.dataset.theme)}catch(e){}toast('Theme saved');themeIcon()};
$('#colbtn').onclick=()=>document.body.classList.toggle('col');
$$('[data-ic]').forEach(e=>e.innerHTML=ic(IC[e.dataset.ic]));
const themeIcon=()=>{const d=root.dataset.theme==='dark'||(!root.dataset.theme&&matchMedia('(prefers-color-scheme:dark)').matches);$('#theme').innerHTML=ic(d?IC.sun:IC.moon)};
buildTypes();buildMerge();buildSplit();buildCompress();themeIcon();
const PAGES=['home','merge','split','compress','topdf','frompdf'],TITLE={home:'Merge & convert',merge:'Merge files',split:'Split PDF',compress:'Compress files',topdf:'Convert to PDF',frompdf:'Convert from PDF'};
const HOME=[['merge','merge','Merge files','Combine several files of the same type into one.',MT.map(m=>[m[1],'#merge/'+m[0]])],['split','split','Split PDF','Cut one PDF into smaller files or pull out pages.',[]],['compress','compress','Compress files','Shrink JPG, PNG, PDF, PowerPoint and Word files to a size you choose.',[]],['topdf','toPdf','Convert to PDF','Turn images, documents and web pages into PDF.',TO.map(t=>[t[1].replace(/ to PDF$/,''),'#topdf/'+t[0]])],['frompdf','fromPdf','Convert from PDF','Get editable files and images out of a PDF.',FROM.map(t=>[t[1].replace(/^PDF to /,''),'#frompdf/'+t[0]])]];
$('#homegrid').replaceChildren(...HOME.map(([k,i,t,d,c])=>h('div',{class:'card hov hc'},h('a',{class:'hl',href:'#'+k},h('span',{class:'hi',html:ic(IC[i])}),h('h3',{},t),h('span',{class:'sm'},d)),c.length?h('div',{class:'hch'},...c.map(([n,u])=>h('a',{href:u},n))):'')));
function route(){const raw=decodeURIComponent(location.hash.slice(1)),[p0,sub]=raw.split('/');
 if(p0==='footer'){$('#footer').scrollIntoView();return}
 const p=PAGES.includes(p0)?p0:'home';
 $$('.page').forEach(s=>s.hidden=s.id!==p);
 $$('#nav a,#side a,.tabs a').forEach(a=>a.classList.toggle('on',a.hash==='#'+p));
 document.title=p==='home'?'Merge & convert: merge, split, compress and convert PDF, Word, Excel, PowerPoint':TITLE[p]+' | Merge & convert';
 if(p==='merge'&&sub&&sub!==mtype&&MT.some(m=>m[0]===sub)){mtype=sub;buildTypes();buildMerge()}
 if(p==='topdf'||p==='frompdf'){const to=p==='topdf',fl=to?'#tflow':'#fflow',t=(to?TO:FROM).find(x=>x[0]===sub);
  if(t){if(OPEN[p]!==sub)openTool(t,fl,to)}else if(OPEN[p]){OPEN[p]=null;$(fl).replaceChildren();markTool(to?'#tg':'#fg',null)}}
 window.scrollTo({top:0,behavior:'instant'})}
addEventListener('hashchange',route);route();
