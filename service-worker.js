const CACHE_NAME='hospital-kpi-v10';
const BASE_URL=self.location.pathname.replace(/service-worker\.js$/,'');
const STATIC_ASSETS=[
  BASE_URL,BASE_URL+'index.html',BASE_URL+'app.js',BASE_URL+'manifest.json',
  'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js'
];
self.addEventListener('install',e=>{
  e.waitUntil(caches.open(CACHE_NAME).then(c=>c.addAll(STATIC_ASSETS)).then(()=>self.skipWaiting()));
});
self.addEventListener('activate',e=>{
  e.waitUntil(caches.keys().then(n=>Promise.all(n.map(c=>c!==CACHE_NAME?caches.delete(c):null))).then(()=>self.clients.claim()));
});
self.addEventListener('fetch',e=>{
  e.respondWith(caches.match(e.request).then(r=>{
    if(r)return r;
    return fetch(e.request).then(n=>{
      if(e.request.method==='GET'&&n.status===200){
        const rc=n.clone();
        caches.open(CACHE_NAME).then(c=>c.put(e.request,rc));
      }
      return n;
    }).catch(()=>{
      if(e.request.mode==='navigate')return caches.match(BASE_URL+'index.html');
      return new Response('Offline',{status:503,headers:{'Content-Type':'text/plain'}});
    });
  }));
});
