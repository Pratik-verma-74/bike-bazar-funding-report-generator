const CACHE_NAME = 'recovery-pdf-v2';

self.addEventListener('install', event => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', event => {
  // Intercept POST request from Web Share Target (WhatsApp / Telegram / Android Share)
  if (event.request.method === 'POST') {
    event.respondWith((async () => {
      try {
        const formData = await event.request.formData();
        let file = formData.get('file');

        // Fallback: Check if file was sent with any other parameter key
        if (!file || typeof file === 'string') {
          for (const [key, value] of formData.entries()) {
            if (value && typeof value === 'object' && value.name) {
              file = value;
              break;
            }
          }
        }

        if (file && typeof file === 'object') {
          const cache = await caches.open('shared-file');
          const cacheResponse = new Response(file, {
            headers: {
              'Content-Type': file.type || 'application/octet-stream',
              'Content-Length': file.size,
              'X-Filename': encodeURIComponent(file.name || 'whatsapp_data.xlsx')
            }
          });
          await cache.put('shared-file-data', cacheResponse);
          // Also cache under /shared-file for backward compatibility
          try {
            await cache.put(new Request('/shared-file'), cacheResponse.clone());
          } catch (e) {}

          return Response.redirect('./?shared=true', 303);
        }
      } catch (err) {
        console.error('Error handling share target in SW:', err);
      }
      return Response.redirect('./', 303);
    })());
    return;
  }

  // Normal fetch handling
  event.respondWith(
    caches.match(event.request).then(response => {
      return response || fetch(event.request);
    }).catch(() => fetch(event.request))
  );
});
