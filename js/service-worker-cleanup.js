// Remove old PWA/service-worker caches that can mix old HTML/CSS/JS with new files,
// then register the current no-cache service worker so the app remains installable.
(async () => {
  const cleanupKey = 'familyAccountsCacheCleanup20260506';
  const registerCurrentServiceWorker = async () => {
    if ('serviceWorker' in navigator) {
      await navigator.serviceWorker.register('./sw.js?v=20260506-1', { scope: './' });
    }
  };

  if (sessionStorage.getItem(cleanupKey) === 'done') {
    await registerCurrentServiceWorker();
    return;
  }

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
    return;
  }
  await registerCurrentServiceWorker();
})().catch(error => {
  console.warn('Service worker cleanup failed', error);
});
