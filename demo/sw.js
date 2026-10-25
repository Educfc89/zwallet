const C='zw-demo-v27';
self.addEventListener('install',e=>{self.skipWaiting();e.waitUntil(caches.open(C).then(c=>c.addAll(['./','./index.html'])).catch(()=>{}))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==C).map(x=>caches.delete(x)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{const r=e.request;if(r.method!=='GET')return;const u=new URL(r.url);
 if(/firestore|identitytoolkit|securetoken|firebaseinstallations/.test(u.hostname))return;
 const ok=u.origin===location.origin||/gstatic\.com|googleapis\.com\/css|cdnjs\.cloudflare|jsdelivr/.test(u.host+u.pathname);if(!ok)return;
 e.respondWith(fetch(r).then(res=>{if(res&&res.status===200){const cp=res.clone();caches.open(C).then(c=>c.put(r,cp)).catch(()=>{});}return res}).catch(()=>caches.match(r).then(m=>m||caches.match('./index.html'))));});
