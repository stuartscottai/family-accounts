const SW_VERSION = 'family-accounts-20260506-1';

self.addEventListener('install', event => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const names = await caches.keys();
    await Promise.all(names.map(name => caches.delete(name)));
    await self.clients.claim();
  })());
});

// Keep a fetch handler for installability without caching stale app files.
self.addEventListener('fetch', event => {
  event.respondWith(fetch(event.request));
});
