// Service Worker for Prayer Tracker PWA - v3 (Offline Support)
const CACHE_NAME = 'prayer-tracker-v3';
const PRECACHE = [
  './',
  './index.html',
  './manifest.json',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(PRECACHE))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  // Skip non-GET requests and ALL Google/API calls
  if (event.request.method !== 'GET' ||
      event.request.url.includes('script.google.com') ||
      event.request.url.includes('googleapis.com') ||
      event.request.url.includes('accounts.google.com') ||
      event.request.url.includes('gstatic.com')) {
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      // Network-first for HTML, cache-first for others
      if (event.request.url.endsWith('.html') || event.request.url.endsWith('/')) {
        return fetch(event.request)
          .then((response) => {
            const clone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
            return response;
          })
          .catch(() => cached || caches.match('./index.html'));
      }
      return cached || fetch(event.request).catch(() => caches.match('./index.html'));
    })
  );
});
