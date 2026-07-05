const CACHE_NAME = 'edueval-pro-v1';
const ASSETS = [
  '/',
  '/login.html',
  '/index.html',
  '/student-form.html',
  '/superadmin.html',
  '/firebase-config.js',
  '/manifest.json'
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS).catch(() => {});
    })
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      );
    })
  );
});

self.addEventListener('fetch', (e) => {
  e.respondWith(
    caches.match(e.request).then((cachedResponse) => {
      return cachedResponse || fetch(e.request).catch(() => {
        // Fallback for offline if page requested is HTML
        if (e.request.headers.get('accept').includes('text/html')) {
          return caches.match('/login.html');
        }
      });
    })
  );
});
