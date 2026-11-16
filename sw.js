const CACHE_NAME="sdda-ledger-pwa-v2";
const CORE_ASSETS=["./","./mobile-v2.html","./app.js","./styles.css","./games.css","./mobile-dashboard.css","./vendor/xlsx.full.min.js","./vendor/pdf-lib.min.js","./manifest.webmanifest","./icons/icon.svg"];
self.addEventListener("install",event=>{event.waitUntil(caches.open(CACHE_NAME).then(cache=>cache.addAll(CORE_ASSETS)).then(()=>self.skipWaiting()));});
self.addEventListener("activate",event=>{event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(key=>key!==CACHE_NAME).map(key=>caches.delete(key)))).then(()=>self.clients.claim()));});
self.addEventListener("fetch",event=>{
 const request=event.request;
 if(request.method!=="GET")return;
 const url=new URL(request.url);
 if(url.origin!==self.location.origin)return;
 if(request.mode==="navigate"){
  event.respondWith(fetch(request).then(response=>{if(response.ok)caches.open(CACHE_NAME).then(cache=>cache.put(request,response.clone()));return response;}).catch(()=>caches.match(request).then(hit=>hit||caches.match("./mobile-v2.html"))));
  return;
 }
 // Always check the network for executable code and styles so a deployed fix
 // cannot be hidden behind a stale versioned cache entry. Keep the last good
 // copy as an offline fallback.
 if(request.destination==="script"||request.destination==="style"){
  event.respondWith(fetch(request).then(response=>{
   if(response.ok)caches.open(CACHE_NAME).then(cache=>cache.put(request,response.clone()));
   return response;
  }).catch(()=>caches.match(request).then(hit=>hit||caches.match(new URL(request.url).pathname.endsWith("/app.js")?"./app.js":request)));
  return;
 }
 event.respondWith(caches.match(request).then(hit=>hit||fetch(request).then(response=>{if(response.ok)caches.open(CACHE_NAME).then(cache=>cache.put(request,response.clone()));return response;})));
});
