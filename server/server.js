// Merge & convert - conversion server. Serves the frontend and /api/convert/:tool.
// Needs on PATH: soffice (LibreOffice), chromium, pdftoppm + pdftotext (poppler), gs (Ghostscript), ocrmypdf (optional, for OCR).
const express=require('express'),multer=require('multer'),cors=require('cors'),{execFile}=require('child_process'),fs=require('fs/promises'),path=require('path'),os=require('os'),JSZip=require('jszip'),{marked}=require('marked'),mammoth=require('mammoth');
const app=express(),up=multer({dest:os.tmpdir(),limits:{fileSize:100*1048576}});
app.use(require('compression')());app.use(cors());app.use(express.static(path.join(__dirname,'..'),{index:'index.html',maxAge:0,etag:true,setHeaders:(r,p)=>{r.setHeader('Cache-Control','no-cache')}}));
const U=require('url').pathToFileURL;
const BIN={soffice:[process.env.SOFFICE_PATH,'soffice','libreoffice','C:\\Program Files\\LibreOffice\\program\\soffice.exe','C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe','/Applications/LibreOffice.app/Contents/MacOS/soffice'],chrome:[process.env.CHROME_PATH,'chromium','chromium-browser','google-chrome','chrome','C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe','C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe','C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe','C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe','/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'],pdftoppm:['pdftoppm'],pdftotext:['pdftotext'],gs:['gs','gswin64c','gswin32c'],ocrmypdf:['ocrmypdf']};
const NEED={soffice:'LibreOffice (libreoffice.org)',chrome:'Google Chrome, Microsoft Edge or Chromium',pdftoppm:'Poppler (pdftoppm)',pdftotext:'Poppler (pdftotext)',gs:'Ghostscript',ocrmypdf:'ocrmypdf with Tesseract'};
const exec1=(c,a)=>new Promise((ok,no)=>execFile(c,a,{timeout:180000,maxBuffer:64*1048576},(e,o,s)=>e?no(e.code==='ENOENT'?Object.assign(Error('x'),{missing:1}):Error(String(s||e.message).slice(0,300))):ok(o)));
const run=async(k,a)=>{for(const c of BIN[k].filter(Boolean)){try{return await exec1(c,a)}catch(e){if(!e.missing)throw e}}throw Error(NEED[k]+' is not installed on this computer, so this conversion cannot run. Install it, or run the app with Docker (see README).')};
const hasCache={};
const has=async k=>{if(hasCache[k]&&Date.now()-hasCache[k].t<60000)return hasCache[k].v;let v=false;for(const c of BIN[k].filter(Boolean)){try{await exec1(c,['--version']);v=true;break}catch(e){if(!e.missing){v=true;break}}}hasCache[k]={t:Date.now(),v};return v};
const winApp=async n=>{if(process.platform!=='win32')return false;try{return /yes/.test(await exec1('powershell',['-NoProfile','-Command',"if(Test-Path 'Registry::HKEY_CLASSES_ROOT\\"+n+".Application'){'yes'}"]))}catch(e){return false}};
const psRun=(script,i,o)=>new Promise((ok,no)=>execFile('powershell',['-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-Command',script],{timeout:240000,env:{...process.env,MC_IN:i,MC_OUT:o}},(e,out,err)=>e?no(Error(String(err||e.message).slice(0,300))):ok(out)));
const COM={Word:"$ErrorActionPreference='Stop';$a=New-Object -ComObject Word.Application;$a.Visible=$false;$a.DisplayAlerts=0;try{$d=$a.Documents.Open($env:MC_IN,$false,$true);$d.SaveAs2($env:MC_OUT,17);$d.Close($false)}finally{$a.Quit()}",PowerPoint:"$ErrorActionPreference='Stop';$a=New-Object -ComObject PowerPoint.Application;try{$p=$a.Presentations.Open($env:MC_IN,-1,0,0);$p.SaveAs($env:MC_OUT,32);$p.Close()}finally{$a.Quit()}"};
// Exact-layout PDF: Microsoft Word / PowerPoint when installed (Windows), otherwise LibreOffice.
const officePdf=async(app,i,d)=>{if(await winApp(app)){try{const o=d+'/out.pdf';await psRun(COM[app],i,o);return o}catch(e){}}
 if(await has('soffice'))return toPdf(i,d);
 throw Error('Exact conversion needs Microsoft '+app+' (Windows) or LibreOffice on the computer that runs this server. Install LibreOffice (free, libreoffice.org) and restart the server.')};
app.get('/api/health',async(q,r)=>{const[soffice,gs,chrome,ocr,word,ppt]=await Promise.all([has('soffice'),has('gs'),has('chrome'),has('ocrmypdf'),winApp('Word'),winApp('PowerPoint')]);r.json({up:true,soffice,gs,chrome,ocr,word,powerpoint:ppt})});
const soffice=(d,a)=>run('soffice',['-env:UserInstallation='+U(d+'/lo').href,'--headless',...a,'--outdir',d]);
const chrome=(i,o,d)=>run('chrome',['--headless','--no-sandbox','--disable-gpu','--no-pdf-header-footer','--user-data-dir='+d+'/chr','--print-to-pdf='+o,U(i).href]);
const THEMES={clean:'body{font:16px/1.6 Arial,sans-serif;max-width:720px;margin:auto}',github:'body{font:16px/1.5 Arial,sans-serif;max-width:760px;margin:auto}code,pre{background:#f6f8fa}h1,h2{border-bottom:1px solid #d0d7de}',academic:'body{font:12pt/1.5 Georgia,serif;max-width:700px;margin:auto;text-align:justify}'};
const toPdf=(i,d)=>soffice(d,['--convert-to','pdf',i]).then(()=>d+'/in.pdf');
const CSS='<meta charset="utf-8"><style>@page{margin:15mm}body{font:12pt/1.5 Calibri,Arial,sans-serif}table{border-collapse:collapse}td,th{border:1px solid #999;padding:4px}img{max-width:100%}</style>';
const SLIDES='<meta charset="utf-8"><style>@page{size:13.33in 7.5in;margin:0}body{font:20pt Arial,sans-serif;margin:0}section{height:7.5in;padding:.6in;box-sizing:border-box;page-break-after:always;overflow:hidden}h1{font-size:32pt;margin:0 0 .3in}</style>';
const htmlPdf=async(html,d)=>{const h=d+'/page.html',o=d+'/out.pdf';await fs.writeFile(h,html);await chrome(h,o,d);return o};
const noSoffice=e=>{if(!/LibreOffice.*not installed/.test(e.message))throw e};
const pptxHtml=async i=>{const z=await JSZip.loadAsync(await fs.readFile(i)),ns=Object.keys(z.files).filter(n=>/^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((a,b)=>a.match(/\d+/)[0]-b.match(/\d+/)[0]);let out='';
 for(const n of ns){const x=await z.file(n).async('string'),ps=[...x.matchAll(/<a:p\b[^>]*>[\s\S]*?<\/a:p>/g)].map(m=>[...m[0].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(t=>t[1]).join('')).filter(Boolean);out+='<section>'+ps.map((p,k)=>k?'<p>'+p+'</p>':'<h1>'+p+'</h1>').join('')+'</section>'}return out};
const H={
 'to-word':async(i,d,o)=>{try{return await officePdf('Word',i,d)}catch(e){if(o.basic&&/Exact conversion/.test(e.message))return htmlPdf(CSS+(await mammoth.convertToHtml({path:i})).value,d);throw e}},
 'to-ppt':(i,d)=>officePdf('PowerPoint',i,d),
 'to-html':async(i,d)=>{const o=d+'/out.pdf';await chrome(i,o,d);return o},
 'to-md':async(i,d,o)=>{const html='<meta charset="utf-8"><style>'+(THEMES[o.theme]||THEMES.clean)+'body{font-size:'+(+o.fs||16)+'px}</style>'+marked.parse(await fs.readFile(i,'utf8'));const h=d+'/page.html',out=d+'/out.pdf';await fs.writeFile(h,html);await chrome(h,out,d);return out},
 'from-jpg':async(i,d,o)=>{await run('pdftoppm',['-jpeg','-r',String(+o.dpi||150),i,d+'/page']);const z=new JSZip();for(const f of(await fs.readdir(d)).filter(f=>/^page.*\.jpg$/.test(f)))z.file(f,await fs.readFile(d+'/'+f));const out=d+'/out.zip';await fs.writeFile(out,await z.generateAsync({type:'nodebuffer'}));return out},
 'from-word':async(i,d,o)=>{let s=i;if(o.ocr==='always'||o.ocr==='auto'&&await has('ocrmypdf')){s=d+'/ocr.pdf';try{await run('ocrmypdf',[o.ocr==='always'?'--force-ocr':'--skip-text','-l',o.lang||'eng',i,s])}catch(e){if(o.ocr==='always')throw e;s=i}}await soffice(d,['--infilter=writer_pdf_import','--convert-to','docx:MS Word 2007 XML',s]);return d+'/'+path.basename(s,'.pdf')+'.docx'},
 'from-ppt':(i,d)=>soffice(d,['--infilter=impress_pdf_import','--convert-to','pptx',i]).then(()=>d+'/in.pptx'),
 'from-pdfa':async(i,d,o)=>{const out=d+'/out.pdf',v=o.ver==='1b'?1:2;
  if(await has('gs')){await run('gs',['-dPDFA='+v,'-dBATCH','-dNOPAUSE','-dQUIET','-sColorConversionStrategy=RGB','-dPDFACompatibilityPolicy=1','-sDEVICE=pdfwrite','-sOutputFile='+out,i]);return out}
  if(await has('soffice')){const od=d+'/pa';await fs.mkdir(od);await run('soffice',['-env:UserInstallation='+U(d+'/lo').href,'--headless','--infilter=draw_pdf_import','--convert-to','pdf:draw_pdf_Export:{"SelectPdfVersion":{"type":"long","value":"'+v+'"}}',i,'--outdir',od]);return od+'/'+path.basename(i,'.pdf')+'.pdf'}
  throw Error('Ghostscript or LibreOffice is not installed on this computer, so this conversion cannot run on the server.')},
 'from-md':async(i,d)=>{const t=await run('pdftotext',['-layout',i,'-']),out=d+'/out.md';await fs.writeFile(out,t.split(/\n{2,}/).map(p=>p.trim()).filter(Boolean).join('\n\n'));return out}};
app.post('/api/convert/:tool',up.single('file'),async(req,res)=>{
 const h=H[req.params.tool];if(!h||!req.file)return res.status(400).json({error:'Unknown tool or missing file.'});
 const d=await fs.mkdtemp(path.join(os.tmpdir(),'mc-')),clean=()=>fs.rm(d,{recursive:true,force:true}).catch(()=>{});
 try{const i=path.join(d,'in'+(path.extname(req.file.originalname).toLowerCase()||'.bin'));await fs.rename(req.file.path,i);
  const out=await h(i,d,JSON.parse(req.body.opts||'{}'));res.sendFile(out,clean)}
 catch(e){res.status(500).json({error:e.message});fs.rm(req.file.path,{force:true}).catch(()=>{});clean()}});
const PORT=process.env.PORT||3000;app.listen(PORT,()=>console.log('Merge & convert running at http://localhost:'+PORT));
