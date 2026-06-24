const CACHE = 'oola-spares-v14';

self.addEventListener('message', e => {
  if (e.data && e.data.type === 'SKIP_WAITING') self.skipWaiting();
});

const STATIC_ASSETS = [
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

// Install — cache static assets only
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE)
      .then(c => c.addAll(STATIC_ASSETS))
      .then(() => self.skipWaiting())
      .catch(err => console.log('[SW] install error:', err))
  );
});

// Activate — clear old caches
self.addEventListener('activate', e => {
  e.waitUntil(
    Promise.all([
      caches.keys().then(keys =>
        Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
      ),
      self.clients.claim()
    ])
  );
});

self.addEventListener('fetch', e => {
  const url = e.request.url;
  const method = e.request.method;

  // Only handle GET
  if (method !== 'GET') return;

  // Google Apps Script API — network only, offline fallback JSON
  if (url.includes('script.google.com')) {
    e.respondWith(
      fetch(e.request.clone())
        .catch(() => new Response(
          JSON.stringify({ success: false, error: 'offline' }),
          { headers: { 'Content-Type': 'application/json' } }
        ))
    );
    return;
  }

  // Google Drive images (lh3.googleusercontent.com) — cache first, network fallback
  if (url.includes('googleusercontent.com') || url.includes('drive.google.com')) {
    e.respondWith(
      caches.match(e.request).then(cached => {
        if (cached) return cached;
        return fetch(e.request.clone())
          .then(res => {
            if (res.ok) {
              const clone = res.clone();
              caches.open(CACHE).then(c => c.put(e.request, clone));
            }
            return res;
          })
          .catch(() => new Response('', { status: 503 }));
      })
    );
    return;
  }

  // Google Fonts — cache first
  if (url.includes('fonts.googleapis.com') || url.includes('fonts.gstatic.com')) {
    e.respondWith(
      caches.match(e.request).then(cached => {
        if (cached) return cached;
        return fetch(e.request.clone())
          .then(res => {
            if (res.ok) caches.open(CACHE).then(c => c.put(e.request, res.clone()));
            return res;
          })
          .catch(() => new Response('', { status: 503 }));
      })
    );
    return;
  }

  // index.html — network first so deploys always update immediately
  if (url.includes('index.html') || url.endsWith('/Spares/') || url.endsWith('/Spares')) {
    e.respondWith(
      fetch(e.request.clone())
        .then(res => {
          if (res && res.status === 200) {
            const clone = res.clone();
            caches.open(CACHE).then(c => c.put(e.request, clone));
          }
          return res;
        })
        .catch(() => caches.match(e.request))
    );
    return;
  }

  // Everything else — cache first, network fallback
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request.clone())
        .then(res => {
          if (res && res.ok) {
            const clone = res.clone();
            caches.open(CACHE).then(c => c.put(e.request, clone));
          }
          return res;
        })
        .catch(() => caches.match('./index.html'));
    })
  );
});
