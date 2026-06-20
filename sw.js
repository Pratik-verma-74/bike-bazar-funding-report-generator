const CACHE_NAME = 'recovery-pdf-v1';

self.addEventListener('install', event => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Intercept POST request from Web Share Target
  if (event.request.method === 'POST' && event.request.url.startsWith(self.registration.scope)) {
    event.respondWith((async () => {
      try {
        const formData = await event.request.formData();
        const file = formData.get('file');

        if (file) {
          const cache = await caches.open('shared-file');
          await cache.put(new Request('/shared-file'), new Response(file, {
            headers: {
              'Content-Type': file.type || 'application/octet-stream',
              'Content-Length': file.size,
              'X-Filename': file.name || 'shared_file.xlsx'
            }
          }));
          return Response.redirect('./?shared=true', 303);
        }
      } catch (err) {
        console.error('Error handling share target:', err);
      }
      return Response.redirect('./', 303);
    })());
    return;
  }

  // Normal fetch handling (basic offline caching)
  event.respondWith(
    caches.match(event.request).then(response => {
      return response || fetch(event.request);
    }).catch(() => fetch(event.request))
  );
});
