/**
 * Аудит СМР — автозапуск + подготовка PWA.
 * Запуск: node start.mjs (или двойной клик по START.bat)
 */
import http from 'node:http';
import path from 'node:path';
import vm from 'node:vm';
import zlib from 'node:zlib';
import { readFile, writeFile, mkdir, stat, copyFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { exec } from 'node:child_process';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const VENDOR = path.join(ROOT, 'vendor');
const ICONS = path.join(ROOT, 'icons');
const SW_VERSION = 'smr-audit-v1.2.0';

async function exists(p){ try{ await stat(p); return true; }catch{ return false; } }

/* ---------- Библиотеки ---------- */
const LIBS=[
  {file:'tailwindcss.js',min:100000,urls:['https://cdn.tailwindcss.com/3.4.17']},
  {file:'xlsx.full.min.js',min:400000,urls:['https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js','https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js']},
  {file:'jszip.min.js',min:50000,urls:['https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js','https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js']}
];
function badCode(code,name){ if(/& &|functio n|ret urn|\bv ar\b/.test(code))return 'файл повреждён'; try{ new vm.Script(code,{filename:name}); return null; }catch(e){ return e.message; } }
async function libOk(lib){ try{ const c=await readFile(path.join(VENDOR,lib.file),'utf8'); return c.length>=lib.min&&!badCode(c,lib.file); }catch{ return false; } }
async function download(url){ const c=new AbortController(); const t=setTimeout(()=>c.abort(),40000); try{ const r=await fetch(url,{signal:c.signal,redirect:'follow'}); if(!r.ok)throw new Error('HTTP '+r.status); return await r.text(); }finally{ clearTimeout(t); } }
async function ensureLibs(){
  console.log('\n=== Шаг 1. Библиотеки ==='); await mkdir(VENDOR,{recursive:true}); const failed=[];
  for(const lib of LIBS){
    process.stdout.write('  '+lib.file+' ... ');
    if(await libOk(lib))console.log('уже на месте ✓');
    else{ let ok=false,err='';
      for(const url of lib.urls){ try{ const code=await download(url); if(code.length<lib.min)throw new Error('файл слишком маленький'); const bad=badCode(code,lib.file); if(bad)throw new Error(bad); await writeFile(path.join(VENDOR,lib.file),code,'utf8'); console.log('скачан ✓'); ok=true; break; }catch(e){ err=e.message; } }
      if(!ok){ console.log('ОШИБКА ✗ — '+err); failed.push(lib); } }
    if(await libOk(lib))await copyFile(path.join(VENDOR,lib.file),path.join(ROOT,lib.file));
  }
  if(failed.length){ console.log('\nПроверьте интернет и запустите снова.'); return false; }
  console.log('  Все библиотеки готовы ✓'); return true;
}

/* ---------- Иконки PNG ---------- */
const CRC_TABLE=(()=>{const t=new Int32Array(256);for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1);t[n]=c;}return t;})();
function crc32(b){let c=0xFFFFFFFF;for(let i=0;i<b.length;i++)c=CRC_TABLE[(c^b[i])&255]^(c>>>8);return (c^0xFFFFFFFF)>>>0;}
function chunk(type,data){const l=Buffer.alloc(4);l.writeUInt32BE(data.length);const t=Buffer.from(type,'ascii');const cr=Buffer.alloc(4);cr.writeUInt32BE(crc32(Buffer.concat([t,data])));return Buffer.concat([l,t,data,cr]);}
function png(size,px){const h=Buffer.alloc(13);h.writeUInt32BE(size,0);h.writeUInt32BE(size,4);h[8]=8;h[9]=6;const stride=size*4;const raw=Buffer.alloc((stride+1)*size);for(let y=0;y<size;y++){raw[y*(stride+1)]=0;px.copy(raw,y*(stride+1)+1,y*stride,(y+1)*stride);}return Buffer.concat([Buffer.from([0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A]),chunk('IHDR',h),chunk('IDAT',zlib.deflateSync(raw,{level:9})),chunk('IEND',Buffer.alloc(0))]);}
const lerp=(a,b,t)=>a+(b-a)*t, C1=[184,58,158], C2=[95,17,82], GREEN=[49,208,140];
function sdRR(x,y,cx,cy,hw,hh,r){const dx=Math.abs(x-cx)-(hw-r),dy=Math.abs(y-cy)-(hh-r);return Math.hypot(Math.max(dx,0),Math.max(dy,0))+Math.min(Math.max(dx,dy),0)-r;}
function segD(px,py,ax,ay,bx,by){const dx=bx-ax,dy=by-ay;const t=Math.max(0,Math.min(1,((px-ax)*dx+(py-ay)*dy)/(dx*dx+dy*dy)));return Math.hypot(px-(ax+dx*t),py-(ay+dy*t));}
function sample(x,y,s,full){
  const t=Math.min(1,Math.max(0,(x+y)/1024)); let r=lerp(C1[0],C2[0],t),g=lerp(C1[1],C2[1],t),b=lerp(C1[2],C2[2],t);
  if((x%64)<1.6||(y%64)<1.6){r=lerp(r,255,.07);g=lerp(g,255,.07);b=lerp(b,255,.07);}
  const X=(x-256)/s+256,Y=(y-256)/s+256;
  const dome=((X-256)**2)/(118*118)+((Y-252)**2)/(98*98)<=1&&Y<=252;
  const brim=sdRR(X,Y,256,262,152,20,20)<=0, ridge=sdRR(X,Y,256,168,20,26,9)<=0;
  if(dome||brim||ridge){ if(brim&&Y>272){r=232;g=232;b=242;}else{r=g=b=255;} }
  const bd=Math.hypot(X-356,Y-352);
  if(bd<=92){r=g=b=255;} if(bd<=80){r=GREEN[0];g=GREEN[1];b=GREEN[2];}
  if(Math.min(segD(X,Y,325,352,347,374),segD(X,Y,347,374,388,322))<=9){r=g=b=255;}
  let a=1; if(!full)a=Math.min(1,Math.max(0,-sdRR(x,y,256,256,256,256,112)+0.5));
  return [r,g,b,a];
}
function renderIcon(size,mask,full){
  const SS=2,total=size*SS,px=Buffer.alloc(size*size*4),s=mask?0.72:1;
  for(let y=0;y<size;y++)for(let x=0;x<size;x++){
    let r=0,g=0,b=0,a=0;
    for(let sy=0;sy<SS;sy++)for(let sx=0;sx<SS;sx++){ const c=sample(((x*SS+sx+.5)/total)*512,((y*SS+sy+.5)/total)*512,s,full); r+=c[0]*c[3];g+=c[1]*c[3];b+=c[2]*c[3];a+=c[3]; }
    const i=(y*size+x)*4; if(a>0){px[i]=Math.round(r/a);px[i+1]=Math.round(g/a);px[i+2]=Math.round(b/a);} px[i+3]=Math.round((a/(SS*SS))*255);
  }
  return png(size,px);
}
const SVG_ANY=`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512"><defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#b83a9e"/><stop offset="1" stop-color="#5f1152"/></linearGradient><pattern id="p" width="64" height="64" patternUnits="userSpaceOnUse"><path d="M64 0H0V64" fill="none" stroke="#fff" stroke-opacity=".07" stroke-width="1.5"/></pattern></defs><rect width="512" height="512" rx="112" fill="url(#g)"/><rect width="512" height="512" rx="112" fill="url(#p)"/><rect x="236" y="142" width="40" height="52" rx="9" fill="#fff"/><path d="M138 252A118 98 0 0 1 374 252Z" fill="#fff"/><rect x="104" y="242" width="304" height="40" rx="20" fill="#fff"/><circle cx="356" cy="352" r="92" fill="#fff"/><circle cx="356" cy="352" r="80" fill="#31D08C"/><path d="M325 352l22 22 41-52" fill="none" stroke="#fff" stroke-width="18" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
const SVG_MASK=`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512"><defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#b83a9e"/><stop offset="1" stop-color="#5f1152"/></linearGradient></defs><rect width="512" height="512" fill="url(#g)"/><g transform="translate(256 256) scale(.72) translate(-256 -256)"><rect x="236" y="142" width="40" height="52" rx="9" fill="#fff"/><path d="M138 252A118 98 0 0 1 374 252Z" fill="#fff"/><rect x="104" y="242" width="304" height="40" rx="20" fill="#fff"/><circle cx="356" cy="352" r="92" fill="#fff"/><circle cx="356" cy="352" r="80" fill="#31D08C"/><path d="M325 352l22 22 41-52" fill="none" stroke="#fff" stroke-width="18" stroke-linecap="round" stroke-linejoin="round"/></g></svg>`;
async function ensureIcons(){
  console.log('\n=== Шаг 2. Иконки ==='); await mkdir(ICONS,{recursive:true});
  for(const [n,c] of [['icon.svg',SVG_ANY],['icon-maskable.svg',SVG_MASK]]){ const p=path.join(ICONS,n); if(!(await exists(p)))await writeFile(p,c,'utf8'); }
  for(const [n,s,m,f] of [['icon-192.png',192,false,false],['icon-512.png',512,false,false],['icon-maskable-512.png',512,true,true],['apple-touch-icon.png',180,false,true]]){
    const p=path.join(ICONS,n); if(!(await exists(p))){ await writeFile(p,renderIcon(s,m,f)); console.log('  '+n+' создан ✓'); } else console.log('  '+n+' уже есть ✓');
  }
}

/* ---------- PWA-файлы ---------- */
const MANIFEST={
  id:'./index.html',name:'Аудит СМР — контроль выполнения работ',short_name:'Аудит СМР',
  description:'Полевой аудит СМР: Excel-матрица, факт прораба, массовые отметки, офлайн.',
  lang:'ru',dir:'ltr',start_url:'./index.html',scope:'./',display:'standalone',orientation:'any',
  background_color:'#2a0725',theme_color:'#951b81',categories:['productivity','business','utilities'],
  icons:[
    {src:'icons/icon-192.png',sizes:'192x192',type:'image/png',purpose:'any'},
    {src:'icons/icon-512.png',sizes:'512x512',type:'image/png',purpose:'any'},
    {src:'icons/icon-maskable-512.png',sizes:'512x512',type:'image/png',purpose:'maskable'},
    {src:'icons/icon.svg',sizes:'any',type:'image/svg+xml',purpose:'any'}
  ]
};
const SW_JS=`/* Аудит СМР — service worker */
const CACHE='${SW_VERSION}';
const CORE=['./','./index.html','./css/styles.css','./js/app.js','./js/pwa.js','./manifest.webmanifest',
 './icons/icon-192.png','./icons/icon-512.png','./icons/icon-maskable-512.png','./icons/apple-touch-icon.png','./icons/icon.svg','./icons/icon-maskable.svg',
 './tailwindcss.js','./xlsx.full.min.js','./jszip.min.js',
 './vendor/tailwindcss.js','./vendor/xlsx.full.min.js','./vendor/jszip.min.js'];
self.addEventListener('install',e=>{e.waitUntil((async()=>{const c=await caches.open(CACHE);await Promise.allSettled(CORE.map(u=>c.add(u)));self.skipWaiting();})());});
self.addEventListener('activate',e=>{e.waitUntil((async()=>{const ks=await caches.keys();await Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)));await self.clients.claim();})());});
self.addEventListener('message',e=>{if(e.data&&e.data.type==='SKIP_WAITING')self.skipWaiting();});
self.addEventListener('fetch',e=>{
  const req=e.request; if(req.method!=='GET')return;
  const url=new URL(req.url); if(url.origin!==self.location.origin)return;
  if(req.mode==='navigate'){e.respondWith((async()=>{try{const f=await fetch(req);const c=await caches.open(CACHE);c.put('./index.html',f.clone());return f;}catch{const c=await caches.open(CACHE);return (await c.match('./index.html'))||(await c.match('./'))||Response.error();}})());return;}
  e.respondWith((async()=>{const c=await caches.open(CACHE);const hit=await c.match(req);if(hit)return hit;try{const f=await fetch(req);if(f&&f.ok)c.put(req,f.clone());return f;}catch{return Response.error();}})());
});
`;
const PWA_JS=`/* Аудит СМР — установка, обновление, сеть */
(function(){
  if(window.__smrPWA)return; window.__smrPWA=true;
  var deferred=null,reg=null,reloading=false;
  function el(id){return document.getElementById(id);}
  function show(e){if(e){e.classList.remove('hidden');e.style.display='flex';}}
  function hide(e){if(e){e.classList.add('hidden');e.style.display='';}}
  window.promptInstall=function(){
    if(deferred){var d=deferred;deferred=null;d.prompt();d.userChoice.then(function(){hide(el('installBtn'));},function(){hide(el('installBtn'));});}
    else alert('Откройте меню браузера (⋮) и выберите «Установить приложение» / «Добавить на главный экран». На iPhone: Safari → Поделиться → На экран «Домой».');
  };
  window.addEventListener('beforeinstallprompt',function(e){e.preventDefault();deferred=e;show(el('installBtn'));});
  window.addEventListener('appinstalled',function(){deferred=null;hide(el('installBtn'));});
  window.applyUpdate=function(){if(reg&&reg.waiting)reg.waiting.postMessage({type:'SKIP_WAITING'});else location.reload();};
  function showUpdate(){show(el('updateBtn'));}
  if('serviceWorker'in navigator&&location.protocol!=='file:'){
    navigator.serviceWorker.register('./sw.js').then(function(r){reg=r;if(r.waiting)showUpdate();
      r.addEventListener('updatefound',function(){var w=r.installing;if(!w)return;w.addEventListener('statechange',function(){if(w.state==='installed'&&navigator.serviceWorker.controller)showUpdate();});});
    }).catch(function(){});
    navigator.serviceWorker.addEventListener('controllerchange',function(){if(!reloading){reloading=true;location.reload();}});
  }
  function net(){var c=el('netChip');if(!c)return;var on=navigator.onLine;c.textContent=on?'Онлайн':'Офлайн';c.style.color=on?'#10b981':'#f59e0b';}
  window.addEventListener('online',net);window.addEventListener('offline',net);net();
})();
`;
async function ensurePWAFiles(){
  console.log('\n=== Шаг 3. Файлы PWA ===');
  await writeFile(path.join(ROOT,'manifest.webmanifest'),JSON.stringify(MANIFEST,null,2),'utf8'); console.log('  manifest.webmanifest ✓');
  await writeFile(path.join(ROOT,'sw.js'),SW_JS,'utf8'); console.log('  sw.js ✓ ('+SW_VERSION+')');
  await mkdir(path.join(ROOT,'js'),{recursive:true});
  await writeFile(path.join(ROOT,'js','pwa.js'),PWA_JS,'utf8'); console.log('  js/pwa.js ✓');
}
async function patchIndex(){
  console.log('\n=== Шаг 4. Проверка index.html ===');
  const p=path.join(ROOT,'index.html'); let html=await readFile(p,'utf8');
  if(!/<html[\s>]/i.test(html)){ console.log('  ВНИМАНИЕ: index.html не похож на HTML. Замените его файлом из поставки.'); return false; }
  const orig=html;
  const inject=t=>{ if(/<\/head>/i.test(html))html=html.replace(/<\/head>/i,t+'\n</head>'); else html=t+'\n'+html; };
  if(!/rel=["']manifest["']/i.test(html))inject('<link rel="manifest" href="manifest.webmanifest">');
  if(!/apple-touch-icon/i.test(html))inject('<link rel="apple-touch-icon" href="icons/apple-touch-icon.png">');
  if(!/name=["']theme-color["']/i.test(html))inject('<meta name="theme-color" content="#951b81">');
  if(!/js\/pwa\.js/i.test(html)){ const t='<script src="js/pwa.js" defer></script>'; if(/<\/body>/i.test(html))html=html.replace(/<\/body>/i,t+'\n</body>'); else html+='\n'+t; }
  if(html!==orig){ await writeFile(p,html,'utf8'); console.log('  Добавлены недостающие подключения ✓'); } else console.log('  Всё подключено ✓');
  return true;
}

/* ---------- Сервер ---------- */
const MIME={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.mjs':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.json':'application/json; charset=utf-8','.webmanifest':'application/manifest+json; charset=utf-8','.svg':'image/svg+xml','.png':'image/png','.jpg':'image/jpeg','.ico':'image/x-icon','.txt':'text/plain; charset=utf-8','.xlsx':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'};
function startServer(){
  const server=http.createServer(async(req,res)=>{
    try{
      if(req.method!=='GET'&&req.method!=='HEAD'){res.writeHead(405);res.end();return;}
      let u=decodeURIComponent((req.url||'/').split('?')[0]); if(u==='/')u='/index.html';
      let full=path.normalize(path.join(ROOT,u));
      if(full!==ROOT&&!full.startsWith(ROOT+path.sep)){res.writeHead(403);res.end('Forbidden');return;}
      let s=null; try{s=await stat(full);}catch{}
      if(s&&s.isDirectory()){full=path.join(full,'index.html');try{s=await stat(full);}catch{s=null;}}
      if(!s){res.writeHead(404,{'Content-Type':'text/plain; charset=utf-8'});res.end('404 — '+u);return;}
      const data=await readFile(full);
      res.writeHead(200,{'Content-Type':MIME[path.extname(full).toLowerCase()]||'application/octet-stream','Content-Length':data.length,'Cache-Control':'no-cache'});
      res.end(req.method==='HEAD'?undefined:data);
    }catch(e){res.writeHead(500,{'Content-Type':'text/plain; charset=utf-8'});res.end('Ошибка: '+e.message);}
  });
  let port=3000;
  server.on('error',e=>{ if(e.code==='EADDRINUSE'&&port<3010){port++;server.listen(port,'127.0.0.1');}else{console.error('Сервер: '+e.message);process.exit(1);} });
  server.on('listening',()=>{
    const url='http://localhost:'+port;
    console.log('\n======================================================');
    console.log('  Аудит СМР '+SW_VERSION+' запущен!');
    console.log('  Адрес: '+url);
    console.log('  НЕ ЗАКРЫВАЙТЕ это окно, пока работаете.');
    console.log('  После проверки заново загрузите ВСЮ папку на хостинг.');
    console.log('======================================================\n');
    const cmd=process.platform==='win32'?'start "" "'+url+'"':process.platform==='darwin'?'open "'+url+'"':'xdg-open "'+url+'"';
    exec(cmd,()=>{});
  });
  server.listen(port,'127.0.0.1');
}

(async()=>{
  console.log('Папка проекта: '+ROOT);
  if(!(await exists(path.join(ROOT,'index.html')))){ console.log('ОШИБКА: рядом нет index.html'); process.exit(1); }
  if(!(await ensureLibs()))process.exit(1);
  await ensureIcons(); await ensurePWAFiles();
  if(!(await patchIndex()))process.exit(1);
  startServer();
})();