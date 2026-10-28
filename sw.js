/* Excel Web — offline cache.
   Cache-first so the app opens with no network at all (aeroplane mode, TWA). */
const CACHE = "excel-web-v1";
const ASSETS = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/maskable-512.png",
];

self.addEventListener("install", e=>{
  e.waitUntil(
    caches.open(CACHE).then(c=>c.addAll(ASSETS).catch(()=>{})).then(()=>self.skipWaiting())
  );
});

self.addEventListener("activate", e=>{
  e.waitUntil(
    caches.keys().then(keys=>Promise.all(
      keys.filter(k=>k!==CACHE).map(k=>caches.delete(k))
    )).then(()=>self.clients.claim())
  );
});

self.addEventListener("fetch", e=>{
  const req = e.request;
  if(req.method !== "GET") return;
  e.respondWith(
    caches.match(req).then(hit=>{
      if(hit) return hit;
      return fetch(req).then(res=>{
        /* only cache same-origin, successful, basic responses */
        if(res && res.status===200 && res.type==="basic"){
          const copy = res.clone();
          caches.open(CACHE).then(c=>c.put(req, copy));
        }
        return res;
      }).catch(()=>caches.match("./index.html"));
    })
  );
});
