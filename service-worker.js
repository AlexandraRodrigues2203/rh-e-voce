const CACHE_NAME = 'rh-e-voce-v8-0-0';
const ASSETS = [
  './',
  './index.html',
  './style.css?v=8',
  './app.js?v=8',
  './config.js?v=8',
  './manifest.json?v=8',
  './apple-touch-icon.png?v=8',
  './favicon-16x16.png?v=8',
  './favicon-32x32.png?v=8',
  './icons/icon-192.png?v=8',
  './icons/icon-512.png?v=8',
  './icons/icon-maskable-512.png?v=8'
];
self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)));
});
self.addEventListener('activate', (event) => {
  event.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  event.respondWith(fetch(event.request).then((response) => {
    const copy = response.clone();
    caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
    return response;
  }).catch(() => caches.match(event.request).then((cached) => cached || caches.match('./index.html'))));
});
