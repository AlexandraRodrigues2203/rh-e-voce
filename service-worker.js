const CACHE_NAME = 'rh-e-voce-v11-20260729';
const ASSETS = [
  './',
  './index.html?v=11',
  './style.css?v=11',
  './app.js?v=11',
  './manifest.json?v=11',
  './logo-rh-e-voce.png?v=11',
  './engrenagem-oficial.png?v=11',
  './apple-touch-icon.png?v=11',
  './favicon-16x16.png?v=11',
  './favicon-32x32.png?v=11',
  './icons/icon-192.png?v=11',
  './icons/icon-512.png?v=11',
  './icons/icon-maskable-512.png?v=11'
];
self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS)));
});
self.addEventListener('activate', event => {
  event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  event.respondWith(fetch(event.request).then(response => {
    const copy = response.clone();
    caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
    return response;
  }).catch(() => caches.match(event.request).then(r => r || caches.match('./'))));
});
