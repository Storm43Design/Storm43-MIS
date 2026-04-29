// Storm43 MIS — Service Worker v1
const CACHE = 'storm43-v1';

// Install
self.addEventListener('install', e => {
  console.log('Storm43 SW installing');
  self.skipWaiting();
});

// Activate
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// Fetch strategy:
// - Google Apps Script = always network (live data)
// - Everything else = network first, fall back to cache
self.addEventListener('fetch', e => {
  const url = e.request.url;

  // Live data — always network, never cache
  if (url.includes('script.google.com') || url.includes('googleapis.com')) {
    e.respondWith(fetch(e.request).catch(() =>
      new Response('{"error":"Offline"}', {headers:{'Content-Type':'application/json'}})
    ));
    return;
  }

  // App shell — network first, cache fallback
  e.respondWith(
    fetch(e.request)
      .then(resp => {
        if (resp && resp.status === 200 && e.request.method === 'GET') {
          const copy = resp.clone();
          caches.open(CACHE).then(c => c.put(e.request, copy));
        }
        return resp;
      })
      .catch(() => caches.match(e.request))
  );
});
