// Remove old PWA/service-worker caches that can mix old HTML/CSS/JS with new files.
(async () => {
  const cleanupKey = 'familyAccountsCacheCleanup20260506';
  if (sessionStorage.getItem(cleanupKey) === 'done') return;

  let changed = false;

  if ('serviceWorker' in navigator) {
    const registrations = await navigator.serviceWorker.getRegistrations();
    await Promise.all(registrations.map(registration => registration.unregister()));
    changed = changed || registrations.length > 0;
  }

  if ('caches' in window) {
    const cacheNames = await caches.keys();
    await Promise.all(cacheNames.map(cacheName => caches.delete(cacheName)));
    changed = changed || cacheNames.length > 0;
  }

  sessionStorage.setItem(cleanupKey, 'done');
  if (changed) {
    window.location.reload();
  }
})().catch(error => {
  console.warn('Service worker cleanup failed', error);
});
