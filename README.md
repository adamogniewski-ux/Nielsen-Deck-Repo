<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Hair Colour · Nielsen</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
:root{--bg1:#e6ecff;--bg2:#fde8f4;--bg3:#e2f7f0;--glass:rgba(255,255,255,.50);--glass2:rgba(255,255,255,.72);--bd:rgba(255,255,255,.7);--hi:rgba(255,255,255,.9);--txt:#1d1d1f;--mut:#6e6e73;--pos:#1fa24a;--neg:#e5303f;--warn:#c77700;--acc:#0071e3;--sh:0 12px 40px rgba(31,38,80,.12);--stick:rgba(244,246,253,.92);--grid:rgba(0,0,0,.06);--blob:.75}
[data-theme=dark]{--bg1:#090b18;--bg2:#190e24;--bg3:#06151a;--glass:rgba(44,46,64,.42);--glass2:rgba(70,72,94,.55);--bd:rgba(255,255,255,.13);--hi:rgba(255,255,255,.25);--txt:#f5f5f7;--mut:#a1a1a6;--pos:#32d74b;--neg:#ff453a;--warn:#ffb340;--acc:#0a84ff;--sh:0 12px 40px rgba(0,0,0,.5);--stick:rgba(30,32,48,.94);--grid:rgba(255,255,255,.07);--blob:.45}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;min-height:100vh;color:var(--txt);font-family:-apple-system,BlinkMacSystemFont,"SF Pro Display","SF Pro Text",Inter,"Segoe UI",system-ui,sans-serif;font-size:14px;letter-spacing:-.011em;font-variant-numeric:tabular-nums;background:linear-gradient(135deg,var(--bg1),var(--bg2) 55%,var(--bg3));background-attachment:fixed;-webkit-font-smoothing:antialiased}
.blob{position:fixed;border-radius:50%;filter:blur(90px);opacity:var(--blob);z-index:-1;animation:drift 24s ease-in-out infinite alternate}
.b1{width:46vw;height:46vw;background:#6f9bff;top:-12vw;left:-10vw}
.b2{width:40vw;height:40vw;background:#ff8fd0;bottom:-14vw;right:-8vw;animation-duration:30s}
.b3{width:32vw;height:32vw;background:#5fe3bd;top:35vh;left:42vw;animation-duration:36s}
@keyframes drift{to{transform:translate(8vw,6vh) scale(1.2)}}
.glass{position:relative;background:var(--glass);backdrop-filter:blur(28px) saturate(190%);-webkit-backdrop-filter:blur(28px) saturate(190%);border:1px solid var(--bd);border-radius:24px;box-shadow:var(--sh),inset 0 1px 0 var(--hi),inset 0 -1px 0 rgba(255,255,255,.06);overflow:hidden}
.glass::before{content:"";position:absolute;inset:0;pointer-events:none;background:radial-gradient(420px circle at var(--mx,50%) var(--my,-30%),rgba(255,255,255,.28),transparent 60%);opacity:.9;transition:opacity .3s}
nav{position:sticky;top:14px;z-index:20;margin:14px auto 0;width:max-content;max-width:calc(100% - 24px);display:flex;gap:10px;align-items:center;padding:7px 8px;border-radius:99px;flex-wrap:wrap;justify-content:center}
.seg{position:relative;display:flex;padding:3px;border-radius:99px;background:rgba(120,120,140,.14)}
.seg button{position:relative;z-index:1;background:none;border:0;padding:7px 18px;border-radius:99px;font-weight:600;color:var(--mut);transition:color .25s}
.seg button.on{color:var(--txt)}
.thumb{position:absolute;top:3px;bottom:3px;border-radius:99px;background:var(--glass2);box-shadow:0 2px 10px rgba(0,0,0,.12),inset 0 1px 0 var(--hi);transition:left .4s cubic-bezier(.3,1.3,.5,1),width .4s cubic-bezier(.3,1.3,.5,1)}
button,select,input,textarea{font:inherit;color:var(--txt)}
.btn{border:1px solid var(--bd);background:var(--glass2);border-radius:99px;padding:8px 16px;font-weight:600;cursor:pointer;transition:transform .2s,background .2s,box-shadow .2s}
.btn:hover{transform:translateY(-1px) scale(1.03);box-shadow:0 6px 18px rgba(0,0,0,.12)}.btn:active{transform:scale(.97)}
.btn.pri{background:var(--acc);color:#fff;border-color:transparent;box-shadow:0 6px 18px rgba(0,113,227,.4)}
.ico{width:36px;height:36px;padding:0;display:grid;place-items:center}
main{max-width:1500px;margin:0 auto;padding:26px 22px 60px}
.hero{display:flex;justify-content:space-between;align-items:flex-end;flex-wrap:wrap;gap:12px;margin:8px 4px 20px}
.hero h1{margin:0;font-size:clamp(28px,4vw,44px);font-weight:700;letter-spacing:-.03em;line-height:1.05}
.hero p{margin:6px 0 0;color:var(--mut);font-size:15px}
.filters{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:20px}
.filters label.t{font-size:12px;color:var(--mut);font-weight:600;text-transform:uppercase;letter-spacing:.06em}
select,input[type=text],input[type=search],input[type=number],input:not([type]){border:1px solid var(--bd);background:var(--glass2);backdrop-filter:blur(14px);border-radius:99px;padding:8px 14px;outline:none;transition:box-shadow .2s}
select:focus,input:focus{box-shadow:0 0 0 3px rgba(0,113,227,.35)}
.chk{display:flex;gap:6px;align-items:center;background:var(--glass2);border:1px solid var(--bd);border-radius:99px;padding:7px 14px;cursor:pointer;font-weight:500}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:16px;margin-bottom:18px}
.kpi{padding:20px 22px}
.kpi .n{font-size:12px;color:var(--mut);font-weight:700;letter-spacing:.07em;text-transform:uppercase}
.kpi .v{font-size:46px;font-weight:700;letter-spacing:-.04em;line-height:1.1;margin:6px 0 2px}
.kpi .v small{font-size:20px;color:var(--mut);font-weight:600}
.kpi .c{font-weight:600}
.kpi .r{display:flex;gap:12px;margin-top:12px;font-size:12px;color:var(--mut);align-items:flex-end;justify-content:space-between}
.kpi svg{display:block}
.pos{color:var(--pos)}.neg{color:var(--neg)}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:18px}
@media(max-width:1000px){.g2{grid-template-columns:1fr}}
.panel{padding:20px 22px}
.panel h3{margin:0 0 12px;font-size:19px;font-weight:700;letter-spacing:-.02em;display:flex;align-items:center;gap:8px}
.cnt{font-size:12px;font-weight:700;background:rgba(120,120,140,.2);border-radius:99px;padding:2px 9px;color:var(--mut)}
.flag{display:flex;gap:12px;align-items:center;padding:11px 12px;margin:0 -12px;border-radius:16px;transition:background .2s,transform .25s}
.flag:hover{background:rgba(120,120,140,.12);transform:translateX(3px)}
.fl-n{flex:1;font-weight:600}.fl-s{color:var(--mut);font-size:12px;font-weight:400;margin-left:4px}
.fl-r{text-align:right;font-weight:600}.fl-r small{display:block;font-weight:500}
.pill{display:inline-block;padding:3px 10px;border-radius:99px;font-size:11px;font-weight:700;margin:5px 5px 0 0;white-space:nowrap;border:1px solid transparent}
.p-neg{background:rgba(255,69,58,.16);color:var(--neg);border-color:rgba(255,69,58,.25)}
.p-pos{background:rgba(48,209,88,.16);color:var(--pos);border-color:rgba(48,209,88,.25)}
.p-warn{background:rgba(255,159,10,.17);color:var(--warn);border-color:rgba(255,159,10,.28)}
.p-info{background:rgba(10,132,255,.15);color:var(--acc);border-color:rgba(10,132,255,.25)}
.in{opacity:0;transform:translateY(16px) scale(.98);animation:in .7s cubic-bezier(.2,.9,.3,1) forwards;animation-delay:calc(var(--i,0)*70ms)}
@keyframes in{to{opacity:1;transform:none}}
.scroll{overflow:auto;max-height:68vh;border-radius:24px}
table{border-collapse:separate;border-spacing:0;width:100%;font-size:13px}
th{position:sticky;top:0;z-index:2;background:var(--stick);padding:12px 12px;text-align:right;font-weight:700;font-size:12px;letter-spacing:.03em;color:var(--mut);cursor:pointer;white-space:nowrap;border-bottom:1px solid var(--grid)}
th:hover{color:var(--txt)}
th:first-child,td:first-child{text-align:left;position:sticky;left:0;background:var(--stick);z-index:1}
th:first-child{z-index:3}
td{padding:8px 12px;text-align:right;white-space:nowrap;border-bottom:1px solid var(--grid);transition:background .15s}
td.fl{text-align:left;white-space:normal;min-width:190px}
tbody tr:hover td{background:rgba(120,120,140,.1)}
tr.sec td{background:rgba(120,120,140,.16)!important;font-weight:700;text-align:left;font-size:12px;letter-spacing:.06em;text-transform:uppercase;color:var(--mut)}
tr.own td:first-child{font-weight:700}
.drop{display:grid;place-items:center;text-align:center;padding:90px 20px;min-height:340px}
.drop b{font-size:30px;letter-spacing:-.03em;display:block;margin-bottom:8px}
.drop span{color:var(--mut);font-size:15px}
.drop .o{width:84px;height:84px;border-radius:26px;margin:0 auto 22px;background:linear-gradient(135deg,#6f9bff,#ff8fd0);box-shadow:0 14px 40px rgba(111,155,255,.5),inset 0 2px 0 rgba(255,255,255,.6);animation:bob 4s ease-in-out infinite;display:grid;place-items:center;font-size:38px}
@keyframes bob{50%{transform:translateY(-10px) rotate(3deg)}}
body.dragging .drop,body.dragging .glass{outline:3px dashed var(--acc);outline-offset:-8px}
#setp{max-height:0;overflow:hidden;opacity:0;transition:max-height .5s,opacity .4s,margin .4s}
#setp.open{max-height:420px;opacity:1;margin-bottom:18px}
.set{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;padding:20px}
.set label{display:flex;flex-direction:column;gap:5px;font-size:12px;color:var(--mut);font-weight:600}
.set .wide{grid-column:1/-1}
dialog{border:1px solid var(--bd);border-radius:26px;width:min(720px,92vw);background:var(--glass2);backdrop-filter:blur(40px) saturate(180%);color:var(--txt);box-shadow:var(--sh);padding:22px}
dialog::backdrop{background:rgba(0,0,0,.35);backdrop-filter:blur(6px)}
textarea{width:100%;height:260px;border:1px solid var(--bd);border-radius:16px;background:rgba(120,120,140,.12);padding:12px;font-family:ui-monospace,Menlo,monospace;font-size:12px}
canvas{max-width:100%}
.note{color:var(--mut);font-size:12px}
</style>
</head>
<body>
<div class="blob b1"></div><div class="blob b2"></div><div class="blob b3"></div>

<nav class="glass">
  <div class="seg" id="tabs"><div class="thumb" id="thumb"></div>
    <button data-t="ov" class="on">Overview</button><button data-t="ex">Explorer</button><button data-t="mk">Markets</button></div>
  <input type="file" id="file" accept=".xlsx,.xls,.csv" hidden>
  <button class="btn pri" onclick="file.click()">Upload</button>
  <button class="btn" onclick="pd.showModal()">Paste</button>
  <button class="btn" onclick="loadDemo()">Demo</button>
  <button class="btn ico" title="Flag settings" onclick="$('setp').classList.toggle('open')">⚙</button>
  <button class="btn ico" id="theme" title="Theme">☾</button>
</nav>

<main>
  <div class="hero in"><div><h1>Hair Colour</h1><p id="sub">Drop a Nielsen file to begin</p></div></div>
  <div id="setp"><div class="glass"><div class="set" id="set"></div></div></div>
  <div id="app"><div class="glass drop in"><div><div class="o">✦</div><b>Drop your Nielsen file</b><span>Excel or CSV · parsed in your browser · flags update instantly</span></div></div></div>
</main>

<dialog id="pd"><h3 style="margin:0 0 12px;font-size:20px">Paste from Excel</h3><textarea id="pt" placeholder="Copy the Nielsen block (including the header rows) and paste here"></textarea>
  <div style="display:flex;gap:10px;margin-top:14px"><button class="btn pri" onclick="loadPaste()">Load</button><button class="btn" onclick="pd.close()">Close</button></div></dialog>

<script>
const $=id=>document.getElementById(id);
const PK=[['y24','2024'],['y25','2025'],['ytd','YTD'],['l13','Latest 13 Wks'],['lp','Latest Period']];
const DEF={minShr:0.3,decl:5,grow:10,minD:25000,swing:10,
 own:"CPD, L OREAL PARIS, GARNIER, PREFERENCE, EXCELLENCE, FERIA, MAGIC ROOT, COLORISTA, CASTING, COLOR GLOSS, AGE PERFECT, COULEUR EXPERTE, NUTRISSE, BELLE COLOR, OLIA, GOOD, COLOR SENSATION"};
let S=Object.assign({},DEF,JSON.parse(localStorage.getItem('hcset')||'{}'));
let M={},tab='ov',sortK=null,sortDir=-1,charts={};
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const norm=s=>String(s).toUpperCase().replace(/[^A-Z0-9 ]/g,'').replace(/\s+/g,' ').trim();
const isOwn=n=>S.own.split(',').map(norm).filter(Boolean).some(o=>norm(n)===o||norm(n).startsWith(o+' '));
const f1=v=>v==null?'–':v.toFixed(1);
const pc=v=>v==null?'–':(v>0?'+':'')+v.toFixed(1)+'%';
const dl=v=>v==null?'–':(v<0?'-':'+')+'$'+Math.abs(Math.round(v)).toLocaleString('en-US');
const cl=v=>v>0?'pos':v<0?'neg':'';
const heat=v=>v==null?'':`background:rgba(${v>0?'48,209,88':'255,69,58'},${Math.min(Math.abs(v)/40,1)*.26})`;

/* ---------- THEME / FX ---------- */
const dark=matchMedia('(prefers-color-scheme:dark)').matches;
document.documentElement.dataset.theme=localStorage.getItem('hctheme')||(dark?'dark':'light');
function themeIcon(){$('theme').textContent=document.documentElement.dataset.theme==='dark'?'☀':'☾'}
$('theme').onclick=()=>{const t=document.documentElement.dataset.theme==='dark'?'light':'dark';document.documentElement.dataset.theme=t;localStorage.setItem('hctheme',t);themeIcon();if(Object.keys(M).length)draw()};
themeIcon();
document.addEventListener('mousemove',e=>{const g=e.target.closest&&e.target.closest('.glass');if(!g)return;const r=g.getBoundingClientRect();g.style.setProperty('--mx',(e.clientX-r.left)+'px');g.style.setProperty('--my',(e.clientY-r.top)+'px')});
function moveThumb(){const b=document.querySelector('#tabs button.on'),t=$('thumb');t.style.left=b.offsetLeft+'px';t.style.width=b.offsetWidth+'px'}
addEventListener('load',moveThumb);addEventListener('resize',moveThumb);
function countUp(){document.querySelectorAll('[data-v]').forEach(el=>{const to=+el.dataset.v,dc=+el.dataset.d||1,t0=performance.now();
 (function f(t){const k=Math.min((t-t0)/900,1),e=1-Math.pow(1-k,4);el.textContent=(to*e).toFixed(dc);if(k<1)requestAnimationFrame(f)})(t0)})}
const css=v=>getComputedStyle(document.documentElement).getPropertyValue(v).trim();
function chartBase(){Chart.defaults.color=css('--mut');Chart.defaults.font.family='-apple-system,"SF Pro Text",Inter,sans-serif';Chart.defaults.borderColor=css('--grid')}

/* ---------- PARSING ---------- */
function per(s){s=s.toLowerCase();
 if(s==='2024')return'y24';if(s==='2025')return'y25';if(/^ytd/.test(s))return'ytd';
 if(/13/.test(s)&&/(latest|last|l13)/.test(s))return'l13';if(/^(latest p|lp|last p|current p)/.test(s))return'lp';return null}
function num(v){if(v==null)return null;let s=String(v).trim();if(!s||s==='-')return null;
 const neg=/^\(.*\)$/.test(s)||/^-/.test(s)||/^\$-/.test(s);s=s.replace(/[^0-9.]/g,'');if(!s)return null;const n=parseFloat(s);return isNaN(n)?null:(neg?-n:n)}
function parseRows(rows,dflt){
 let cur=null,sec=null,blank=false,pend=null,out=[];
 for(const r of rows){
  const c=r.map(x=>x==null?'':String(x).trim());
  const hits={};c.forEach((x,i)=>{const k=per(x);if(k&&!(k in hits))hits[k]=i});
  const mi=c.findIndex(x=>/^market$/i.test(x));
  if(mi>=0){const nm=c.slice(mi+1).find(x=>x&&!per(x));if(nm)pend=nm}
  if(Object.keys(hits).length>=3){
    cur={name:pend||dflt,starts:hits,labels:Object.fromEntries(c.map((x,i)=>[per(x),x]).filter(a=>a[0])),sections:[]};
    sec={name:'Brand & franchise',rows:[]};cur.sections.push(sec);blank=false;pend=null;out.push(cur);continue}
  if(!cur)continue;
  if(c.every(x=>x==='')){blank=true;continue}
  const name=c[0];if(!name||name==='$/Units')continue;
  const p={};let any=false;
  for(const [k] of PK){const st=cur.starts[k];if(st==null)continue;
    const o={shr:num(c[st]),chg:num(c[st+1]),d:num(c[st+2])};
    if(o.shr!=null||o.d!=null){any=true}p[k]=o}
  if(!any)continue;
  if(blank){sec={name:/^TL /i.test(name)?name:'Other',rows:[]};cur.sections.push(sec);blank=false}
  sec.rows.push({name,sec:sec.name,p})
 }
 return out}
function ingest(blocks){M={};blocks.forEach(b=>{let n=b.name,i=2;while(M[n])n=b.name+' ('+i++ +')';M[n]=b});
 if(!Object.keys(M).length){alert('Could not find a Nielsen header row (2024 / 2025 / YTD / Latest 13 Wks / Latest Period).');return}
 const b=Object.values(M)[0];$('sub').textContent=Object.keys(M).length+' market(s) · '+(b.labels.ytd||'YTD')+' · '+(b.labels.lp||'Latest Period');
 init()}
function handleFile(f){if(!f)return;const rd=new FileReader();
 rd.onload=ev=>{const wb=XLSX.read(ev.target.result,{type:'array'});let all=[];
  wb.SheetNames.forEach(n=>{const rows=XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,raw:false,defval:''});all=all.concat(parseRows(rows,n))});
  ingest(all)};rd.readAsArrayBuffer(f)}
$('file').onchange=e=>handleFile(e.target.files[0]);
addEventListener('dragover',e=>{e.preventDefault();document.body.classList.add('dragging')});
addEventListener('dragleave',e=>{if(!e.relatedTarget)document.body.classList.remove('dragging')});
addEventListener('drop',e=>{e.preventDefault();document.body.classList.remove('dragging');handleFile(e.dataTransfer.files[0])});
function loadPaste(){const rows=$('pt').value.split(/\r?\n/).map(l=>l.split('\t'));pd.close();ingest(parseRows(rows,'Pasted market'))}

/* ---------- FLAGS ---------- */
function flags(r,pk){
 const p=r.p[pk]||{},a=r.p.ytd||{},b=r.p.l13||{},o=[];
 const big=(p.shr||0)>=S.minShr||Math.abs(p.d||0)>=S.minD*2;if(!big||p.chg==null&&p.d==null&&pk!=='lp')return o;
 if(p.chg!=null&&p.d!=null){
  if(p.chg<=-S.decl&&p.d<=-S.minD)o.push({c:'neg',t:'Losing'});
  if(p.chg>=S.grow&&p.d>=S.minD)o.push({c:'pos',t:'Surging'})}
 if(pk==='lp'||pk==='l13'){
  if(a.chg!=null&&p.chg!=null){if(a.chg>0&&p.chg<0&&Math.abs(p.d||0)>=S.minD/2)o.push({c:'warn',t:'Turned negative'});
   if(a.chg<0&&p.chg>0&&Math.abs(p.d||0)>=S.minD/2)o.push({c:'pos',t:'Turned positive'})}
  const base=pk==='lp'?b:a;if(base.chg!=null&&p.chg!=null&&Math.abs(p.d||0)>=S.minD/2){
   const dd=p.chg-base.chg;if(dd>=S.swing)o.push({c:'info',t:'Accelerating'});if(dd<=-S.swing)o.push({c:'warn',t:'Decelerating'})}}
 if(!(r.p.y25&&r.p.y25.shr)&&p.shr>0&&r.p.ytd)o.push({c:'info',t:'New'});
 return o}
const pill=f=>`<span class="pill p-${f.c}">${f.t}</span>`;
const AGG=new Set(['MASS','CPD','L OREAL PARIS','GARNIER','WELLA','CLAIROL','HENKEL'].map(norm));
const ARR=m=>M[m].sections.flatMap(s=>s.rows);

/* ---------- UI ---------- */
function init(){
 $('app').innerHTML=`<div class="filters in" id="bar"></div><div id="view"></div>`;
 const ms=Object.keys(M);
 $('bar').innerHTML=`<label class="t">Market</label><select id="mk" onchange="draw()">${ms.map(m=>`<option>${esc(m)}</option>`).join('')}</select>
 <label class="t">Period</label><select id="pk" onchange="draw()">${PK.slice(2).reverse().map(([k,l])=>`<option value="${k}">${l}</option>`).join('')}</select>
 <label class="chk"><input type="checkbox" id="ownOnly" checked onchange="draw()"> Own brands</label>`;
 draw()}
function sets(){const f=[['minShr','Min share (pts)'],['decl','Decline % ≥'],['grow','Growth % ≥'],['minD','Min $ change'],['swing','Swing (pts)']];
 $('set').innerHTML=f.map(([k,l])=>`<label>${l}<input type="number" step="any" value="${S[k]}" onchange="S['${k}']=+this.value;save()"></label>`).join('')+
 `<label class="wide">Own-brand list (comma separated)<input id="own" value="${esc(S.own)}" onchange="S.own=this.value;save()"></label>`}
function save(){localStorage.setItem('hcset',JSON.stringify(S));if(Object.keys(M).length)draw()}
document.querySelectorAll('#tabs button').forEach(b=>b.onclick=()=>{tab=b.dataset.t;document.querySelectorAll('#tabs button').forEach(x=>x.classList.toggle('on',x===b));moveThumb();if(Object.keys(M).length)draw()});
function draw(){const m=$('mk').value,pk=$('pk').value;Object.values(charts).forEach(c=>c.destroy());charts={};chartBase();
 ({ov:overview,ex:explorer,mk:compare}[tab])(m,pk)}

function spark(r){const ks=['y25','ytd','l13','lp'],v=ks.map(k=>(r.p[k]||{}).chg),mx=Math.max(1,...v.map(x=>Math.abs(x||0)));
 return`<svg width="${ks.length*16}" height="34">${v.map((x,i)=>{const h=Math.max(2,Math.abs(x||0)/mx*15);return`<rect x="${i*16}" y="${x>=0?17-h:17}" width="11" height="${h}" rx="3" fill="${(x||0)>=0?css('--pos'):css('--neg')}" opacity="${.45+.55*(i+1)/ks.length}"/>`}).join('')}</svg>`}

function overview(m,pk){
 const rows=ARR(m),lab=M[m].labels,first=M[m].sections[0].rows;
 const pick=n=>first.find(r=>norm(r.name)===norm(n))||rows.find(r=>norm(r.name)===norm(n));
 const cards=['CPD','L OREAL PARIS','GARNIER','CLAIROL'].map(pick).filter(Boolean).map((r,i)=>{const p=r.p[pk]||{};
  return`<div class="glass kpi in" style="--i:${i}"><div class="n">${esc(r.name)}</div><div class="v"><span data-v="${p.shr??0}" data-d="1">0</span></div>
  <div class="c ${cl(p.chg)}">${pc(p.chg)} <span style="color:var(--mut);font-weight:500">vs YA · ${dl(p.d)}</span></div>
  <div class="r"><div>${spark(r)}<div style="margin-top:2px">25 · YTD · 13W · LP</div></div></div></div>`}).join('');
 const own=$('ownOnly').checked,seen=new Set(),F=[];
 rows.forEach(r=>{const k=r.sec+'|'+norm(r.name);if(seen.has(k))return;seen.add(k);if(own&&!isOwn(r.name))return;
  const f=flags(r,pk);if(f.length)F.push({r,f})});
 const seenN=new Set(),FF=F.filter(x=>{const k=norm(x.r.name)+x.f.map(y=>y.t).join();if(seenN.has(k))return false;seenN.add(k);return true})
  .sort((a,b)=>Math.abs((b.r.p[pk]||{}).d||0)-Math.abs((a.r.p[pk]||{}).d||0));
 const item=x=>{const p=x.r.p[pk]||{};return`<div class="flag"><div class="fl-n">${esc(x.r.name)}<span class="fl-s">${esc(x.r.sec)}</span><div>${x.f.map(pill).join('')}</div></div>
  <div class="fl-r">${f1(p.shr)} <span style="color:var(--mut);font-weight:500">shr</span><small class="${cl(p.chg)}">${pc(p.chg)} · ${dl(p.d)}</small></div></div>`};
 const bad=FF.filter(x=>x.f.some(y=>y.c==='neg'||y.c==='warn')),good=FF.filter(x=>!bad.includes(x));
 const mv=[];const sn=new Set();rows.forEach(r=>{const n=norm(r.name);if(sn.has(n)||AGG.has(n)||/^TL /.test(n))return;sn.add(n);if(own&&!isOwn(r.name))return;const d=(r.p[pk]||{}).d;if(d!=null)mv.push([r.name,d])});
 mv.sort((a,b)=>b[1]-a[1]);const top=[...mv.slice(0,8),...mv.slice(-8)].filter((x,i,a)=>a.indexOf(x)===i);
 $('view').innerHTML=`<div class="kpis">${cards}</div>
 <div class="g2"><div class="glass panel in" style="--i:4"><h3>Needs attention <span class="cnt">${bad.length}</span></h3>${bad.slice(0,12).map(item).join('')||'<div class="note">Nothing flagged 🎉</div>'}</div>
 <div class="glass panel in" style="--i:5"><h3>Momentum &amp; wins <span class="cnt">${good.length}</span></h3>${good.slice(0,12).map(item).join('')||'<div class="note">Nothing flagged</div>'}</div></div>
 <div class="glass panel in" style="--i:6"><h3>Top $ movers vs YA · ${esc(lab[pk]||pk)}</h3><canvas id="ch" height="90"></canvas></div>`;
 countUp();
 const pos=css('--pos'),neg=css('--neg');
 charts.a=new Chart($('ch'),{type:'bar',data:{labels:top.map(x=>x[0]),datasets:[{data:top.map(x=>x[1]),backgroundColor:top.map(x=>x[1]>=0?pos:neg),borderRadius:10,borderSkipped:false}]},
  options:{animation:{duration:1100,easing:'easeOutQuart'},plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>dl(c.raw)}}},scales:{y:{ticks:{callback:v=>dl(v)},grid:{color:css('--grid')}},x:{grid:{display:false},ticks:{maxRotation:60}}}}})}

function explorer(m,pk){
 const secs=M[m].sections;
 $('view').innerHTML=`<div class="filters in"><label class="t">Section</label><select id="sec" onchange="tbl()"><option value="">All</option>${secs.map(s=>`<option>${esc(s.name)}</option>`).join('')}</select>
 <label class="t">Metric</label><select id="met" onchange="tbl()"><option value="chg">% Chg YA</option><option value="shr">$ Share</option><option value="d">$ Chg YA</option></select>
 <input type="search" id="q" placeholder="Search franchise…" oninput="tbl()"><label class="chk"><input type="checkbox" id="fo" onchange="tbl()"> Flagged only</label>
 <button class="btn" onclick="csv()">Export CSV</button></div><div class="glass in" style="--i:1"><div class="scroll"><table id="t"></table></div></div>`;tbl()}
function tbl(){const m=$('mk').value,pk=$('pk').value,met=$('met').value,q=$('q').value.toLowerCase(),sv=$('sec').value,fo=$('fo').checked,lab=M[m].labels;
 const fm=met==='chg'?pc:met==='shr'?f1:dl;
 let h=`<thead><tr><th>Franchise</th>${PK.map(([k,l])=>`<th onclick="sortK='${k}';sortDir*=-1;tbl()">${esc(lab[k]||l)}${sortK===k?(sortDir<0?' ↓':' ↑'):''}</th>`).join('')}<th style="text-align:left">Flags · ${esc(lab[pk]||pk)}</th></tr></thead><tbody>`;
 M[m].sections.forEach(s=>{if(sv&&s.name!==sv)return;let rs=s.rows.filter(r=>(!q||r.name.toLowerCase().includes(q)));
  rs=rs.map(r=>({r,f:flags(r,pk)})).filter(x=>!fo||x.f.length);
  if(sortK)rs.sort((a,b)=>sortDir*(((a.r.p[sortK]||{})[met]??-1e15)-((b.r.p[sortK]||{})[met]??-1e15))*-1);
  if(!rs.length)return;h+=`<tr class="sec"><td colspan="${PK.length+2}">${esc(s.name)}</td></tr>`;
  rs.forEach(({r,f})=>{h+=`<tr class="${isOwn(r.name)?'own':''}"><td>${esc(r.name)}</td>${PK.map(([k])=>{const v=(r.p[k]||{})[met];
   return`<td class="${met==='shr'?'':cl(v)}" style="${met==='shr'?'':heat((r.p[k]||{}).chg)}">${fm(v)}</td>`}).join('')}<td class="fl">${f.map(pill).join('')}</td></tr>`})});
 $('t').innerHTML=h+'</tbody>'}
function csv(){const m=$('mk').value,pk=$('pk').value;const o=[['Market','Section','Franchise',...PK.flatMap(([k,l])=>[l+' Shr',l+' %Chg',l+' $Chg']),'Flags']];
 M[m].sections.forEach(s=>s.rows.forEach(r=>o.push([m,s.name,r.name,...PK.flatMap(([k])=>{const p=r.p[k]||{};return[p.shr??'',p.chg??'',p.d??'']}),flags(r,pk).map(x=>x.t).join('; ')])));
 const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([o.map(r=>r.map(x=>`"${String(x).replace(/"/g,'""')}"`).join(',')).join('\n')],{type:'text/csv'}));a.download='nielsen_flags.csv';a.click()}

function compare(m,pk){
 const names=[...new Set(Object.values(M).flatMap(b=>b.sections.flatMap(s=>s.rows.map(r=>r.name))))];
 $('view').innerHTML=`<div class="filters in"><label class="t">Franchise</label><input id="it" list="dl" value="${esc(names.find(n=>norm(n)==='CPD')||names[0])}" onchange="cmp()"><datalist id="dl">${names.map(n=>`<option value="${esc(n)}">`).join('')}</datalist></div>
 <div class="g2"><div class="glass panel in" style="--i:1"><h3>% Chg YA by market</h3><canvas id="ch" height="170"></canvas></div><div class="glass panel in" style="--i:2"><h3>Detail</h3><div id="ct"></div></div></div>`;cmp()}
function cmp(){const pk=$('pk').value,n=norm($('it').value);charts.c&&charts.c.destroy();
 const rs=Object.entries(M).map(([mn,b])=>{const r=b.sections.flatMap(s=>s.rows).find(r=>norm(r.name)===n);return r?{m:mn,p:r.p[pk]||{}}:null}).filter(Boolean);
 if(!rs.length){$('ct').innerHTML='<div class="note">Not found.</div>';return}
 const base=rs[0],gap=x=>x===base||x.p.shr==null||base.p.shr==null?'–':((x.p.shr-base.p.shr>0?'+':'')+(x.p.shr-base.p.shr).toFixed(1));
 $('ct').innerHTML=`<table><thead><tr><th>Market</th><th>Share</th><th>Gap vs 1st</th><th>% Chg</th><th>$ Chg</th></tr></thead><tbody>${rs.map(x=>`<tr><td>${esc(x.m)}</td><td>${f1(x.p.shr)}</td><td>${gap(x)}</td><td class="${cl(x.p.chg)}">${pc(x.p.chg)}</td><td class="${cl(x.p.d)}">${dl(x.p.d)}</td></tr>`).join('')}</tbody></table>`;
 charts.c=new Chart($('ch'),{type:'bar',data:{labels:rs.map(x=>x.m),datasets:[{data:rs.map(x=>x.p.chg),backgroundColor:rs.map(x=>(x.p.chg||0)>=0?css('--pos'):css('--neg')),borderRadius:12,borderSkipped:false}]},options:{animation:{duration:900},plugins:{legend:{display:false}},scales:{y:{ticks:{callback:v=>v+'%'},grid:{color:css('--grid')}},x:{grid:{display:false}}}}})}

/* ---------- DEMO ---------- */
const DEMO=`MARKET	NATIONAL EX NFLD	2024			2025			YTD			Latest 13 Wks			Latest Period		
$/Units	$	$ Shr - Product	$ % Chg YA	$ Chg YA	$ Shr - Product	$ % Chg YA	$ Chg YA	$ Shr - Product	$ % Chg YA	$ Chg YA	$ Shr - Product	$ % Chg YA	$ Chg YA	$ Shr - Product	$ % Chg YA	$ Chg YA
MASS		100.0	0.0%	 $213,234,923 	100.0	2.6%	 $5,525,401 	100.0	3.3%	 $6,102,134 	100.0	3.2%	 $2,001,101 	100.0	2.5%	 $610,443 
CPD		61.8	0.0%	 $131,693,406 	60.8	1.0%	 $1,377,156 	53.9	2.1%	 $2,128,100 	53.5	1.8%	 $617,128 	53.6	1.9%	 $247,240 
L OREAL PARIS		39.4	0.0%	 $84,052,665 	38.4	-0.1%	-$81,664 	33.8	1.6%	 $1,017,891 	33.7	1.5%	 $316,058 	33.5	0.7%	 $58,782 
PREFERENCE		13.4	0.0%	 $28,541,232 	11.4	-12.7%	-$3,615,702 	9.1	-9.4%	-$1,789,897 	8.6	-11.8%	-$750,354 	8.1	-16.5%	-$400,428 
EXCELLENCE		9.5	0.0%	 $20,358,354 	10.5	13.3%	 $2,716,458 	9.5	4.9%	 $838,008 	9.8	4.2%	 $258,660 	10.1	3.7%	 $90,907 
FERIA		6.7	0.0%	 $14,231,038 	6.6	0.8%	 $120,058 	6.1	8.1%	 $865,556 	6.0	6.9%	 $254,220 	6.1	8.4%	 $117,449 
MAGIC ROOT COVER UP		4.2	0.0%	 $8,942,851 	4.1	0.9%	 $82,476 	3.3	-5.3%	-$356,413 	3.1	-6.7%	-$146,905 	3.2	-2.9%	-$23,487 
CASTING		0.9	0.0%	 $1,894,283 	0.8	-2.7%	-$51,350 	1.3	74.9%	 $1,060,358 	1.7	134.5%	 $619,903 	1.7	149.8%	 $261,106 
GARNIER		22.3	0.0%	 $47,638,396 	22.4	3.1%	 $1,458,930 	20.1	3.3%	 $1,216,113 	19.8	2.5%	 $319,547 	20.1	4.0%	 $194,120 
NUTRISSE		11.3	0.0%	 $24,029,173 	10.3	-6.3%	-$1,506,089 	8.8	-2.0%	-$330,963 	8.5	-0.9%	-$50,511 	8.5	1.2%	 $24,555 
OLIA		4.3	0.0%	 $9,115,372 	4.1	-1.1%	-$104,546 	3.5	-3.7%	-$253,911 	3.5	-3.7%	-$87,706 	3.5	-6.7%	-$63,357 
GOOD				  	0.9	0.0%	 $2,045,084 	1.1	40.6%	 $584,427 	1.0	14.3%	 $84,166 	1.0	13.8%	 $31,557 
CLAIROL		19.5	0.0%	 $41,677,641 	19.4	1.9%	 $794,347 	17.3	3.3%	 $1,039,789 	17.2	2.3%	 $255,886 	17.1	-1.0%	-$43,858 
NATURAL INSTINCTS		5.0	0.0%	 $10,755,901 	5.0	1.2%	 $127,766 	4.3	1.3%	 $105,772 	4.2	-3.6%	-$102,338 	4.0	-13.7%	-$159,919 
MANIC PANIC		2.4	0.0%	 $5,163,230 	3.1	33.0%	 $1,701,533 	3.6	37.3%	 $1,850,143 	4.0	31.0%	 $616,690 	3.9	29.1%	 $222,352 
REVLON		0.8	0.0%	 $1,632,843 	0.7	-8.1%	-$132,560 	0.9	49.3%	 $559,215 	1.0	69.9%	 $272,710 	1.0	59.0%	 $94,675 
`;
function loadDemo(){ingest(parseRows(DEMO.split('\n').map(l=>l.split('\t')),'Demo'))}
sets();
</script>
</body>
</html>
# Nielsen-Deck-Repo
Dynamic Nielsen Hair Colour Deck
