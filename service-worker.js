const CACHE = "rh-e-voce-v9-20260729";
const ASSETS = [
  "./",
  "./index.html",
  "./style.css?v=9",
  "./app.js?v=9",
  "./config.js?v=9",
  "./manifest.json?v=9",
  "./logo-rh-e-voce.png?v=9",
  "./apple-touch-icon.png?v=9",
  "./icons/icon-192.png?v=9",
  "./icons/icon-512.png?v=9",
  "./icons/icon-maskable-512.png?v=9"
];
self.addEventListener("install", e => {e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));self.skipWaiting();});
self.addEventListener("activate", e => {e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))));self.clients.claim();});
self.addEventListener("fetch", e => {
  if (e.request.method !== "GET") return;
  e.respondWith(fetch(e.request).then(r => {const copy=r.clone();caches.open(CACHE).then(c=>c.put(e.request,copy));return r;}).catch(()=>caches.match(e.request).then(r=>r||caches.match("./index.html"))));
});
