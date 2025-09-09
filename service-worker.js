// ---- Family Accounts PWA (single-file site) ----
const BASE = '/family-accounts/';
const CACHE = 'fa-shell-v1';

// Your “shell” is just the HTML + manifest + icons
const SHELL = [
  BASE,
  BASE + 'index.html',
  BASE + 'manifest.webmanifest',
  BASE + 'icon-192.png',
  BASE + 'icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

// Cache-first for our own static files under /family-accounts/;
// let Google Drive requests go to the network.
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  const sameOrigin = url.origin === self.location.origin;
  const underBase = sameOrigin && url.pathname.startsWith(BASE);

  if (underBase) {
    event.respondWith(
      caches.match(event.request).then(r => r || fetch(event.request))
    );
  }
  // For Drive/API requests, do nothing special (network).
});
