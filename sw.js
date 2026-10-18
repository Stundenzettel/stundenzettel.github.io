// Hält die App offline verfügbar. Bei jeder neuen Version CACHE hochzählen.
const CACHE = 'stundenzettel-v12';
const ASSETS = [
  './',
  './index.html',
  './app.css',
  './app.js',
  './xlsx.js',
  './pdf.js',
  './vorlage.xlsx',
  './manifest.webmanifest',
  './icons/icon-180.png',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE)
      .then(cache => cache.addAll(ASSETS.map(url => new Request(url, { cache: 'reload' }))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Zuerst aus dem Cache (funktioniert ohne Netz), im Hintergrund aktualisieren.
self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  event.respondWith(
    caches.open(CACHE).then(cache =>
      cache.match(req, { ignoreSearch: true }).then(hit => {
        const update = fetch(req)
          .then(res => { if (res.ok) cache.put(req, res.clone()); return res; })
          .catch(() => hit);
        if (hit) { event.waitUntil(update); return hit; }
        return update;
      })
    )
  );
});
