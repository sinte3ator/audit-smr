/* Аудит СМР — service worker */
const CACHE='smr-audit-v1.2.0';
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
