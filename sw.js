const CACHE_NAME = 'nofly-sedes-v23';
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './guia_campo.html',
  './analise_frequencia.html',
  './relatorio.html',
  './manual.html',
  './dados_mapa.js',
  './vistorias_kobo.json',
  './manifest.json',
  './favicon.ico',
  './apple-touch-icon.png',
  './icon-192.png',
  './icon-512.png',
  './logos/favicon.ico',
  './logos/icon-192.png',
  './logos/icon-512.png',
  './logos/apple-touch-icon.png',
  './logos/NoFly.jpg',
  './logos/Logo_sede.jpg',
  './logos/logo clealco.png'
];

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS_TO_CACHE).catch(() => {});
    })
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((name) => {
          if (name !== CACHE_NAME) {
            return caches.delete(name);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    fetch(event.request)
      .then((networkResponse) => {
        if (networkResponse && networkResponse.status === 200 && networkResponse.type === 'basic') {
          const responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseToCache);
          });
        }
        return networkResponse;
      })
      .catch(() => {
        return caches.match(event.request);
      })
  );
});