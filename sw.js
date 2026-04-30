// ═══════════════════════════════════════════════════════════════
// Vale Digital — Service Worker v4.0
// Estrategia: cache-first para assets propios
// GAS: siempre red directa, sin intercepción del SW
// ═══════════════════════════════════════════════════════════════

const CACHE_NAME = 'vale-digital-v5.0';

const ASSETS_ESTATICOS = [
  '/vale/',
  '/vale/index.html',
  '/vale/manifest.json',
  '/vale/icon-192.png',
  '/vale/icon-512.png',
];

// ── INSTALL ─────────────────────────────────────────────────────
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(ASSETS_ESTATICOS))
      .then(() => self.skipWaiting())
  );
});

// ── ACTIVATE: limpiar cachés viejas ─────────────────────────────
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(k => k !== CACHE_NAME)
          .map(k => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

// ── FETCH ────────────────────────────────────────────────────────
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // GAS → red directa, NUNCA interceptar (GET ni POST)
  if (url.hostname === 'script.google.com') {
    return; // el browser maneja directo sin SW
  }

  // Google Identity Services → red directa
  if (url.hostname === 'accounts.google.com') {
    return;
  }

  // Quill CDN → cache-first
  if (url.hostname === 'cdn.quilljs.com') {
    event.respondWith(cacheFirst(event.request));
    return;
  }

  // Assets propios → cache-first
  if (url.origin === self.location.origin) {
    event.respondWith(cacheFirst(event.request));
    return;
  }
});

// ── CACHE-FIRST ──────────────────────────────────────────────────
async function cacheFirst(request) {
  const cached = await caches.match(request);
  if (cached) return cached;
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_NAME);
      cache.put(request, response.clone());
    }
    return response;
  } catch (_) {
    return new Response('', { status: 503, statusText: 'Offline' });
  }
}
