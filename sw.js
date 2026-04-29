// ═══════════════════════════════════════════════════════════════
// Vale Digital — Service Worker v2.0
// Estrategia: cache-first para assets, network-first para GAS
// ═══════════════════════════════════════════════════════════════

const CACHE_NAME = 'vale-digital-v2.0';

// Assets que se cachean en la instalación
const ASSETS_ESTATICOS = [
  '/vale/',
  '/vale/index.html',
  '/vale/manifest.json',
  '/vale/icon-192.png',
  '/vale/icon-512.png',
  // Quill CDN — se cachea en el primer uso (runtime)
];

// ── INSTALL: pre-cachear assets propios ─────────────────────────
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

// ── FETCH: lógica de caché ───────────────────────────────────────
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // GAS (script.google.com) → network-first (sin cachear POSTs)
  if (url.hostname === 'script.google.com') {
    event.respondWith(networkFirst(event.request));
    return;
  }

  // Quill CDN (cdn.quilljs.com) → cache-first con fallback de red
  if (url.hostname === 'cdn.quilljs.com') {
    event.respondWith(cacheFirst(event.request));
    return;
  }

  // Google Identity Services → solo red (no cachear tokens)
  if (url.hostname === 'accounts.google.com') {
    event.respondWith(fetch(event.request));
    return;
  }

  // Assets propios (index.html, manifest, íconos) → cache-first
  if (url.origin === self.location.origin) {
    event.respondWith(cacheFirst(event.request));
    return;
  }

  // Todo lo demás → red directa
  event.respondWith(fetch(event.request));
});

// ── ESTRATEGIAS ──────────────────────────────────────────────────

// Cache-first: sirve desde caché, si no hay va a la red y guarda
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
    // Sin red y sin caché → respuesta vacía (no rompe la app)
    return new Response('', { status: 503, statusText: 'Offline' });
  }
}

// Network-first: intenta la red, si falla usa caché
// IMPORTANTE: los requests POST no se cachean (Cache API no lo soporta)
async function networkFirst(request) {
  const esPost = request.method === 'POST';
  try {
    const response = await fetch(request);
    // Solo cachear GETs — nunca POSTs
    if (response.ok && !esPost) {
      const cache = await caches.open(CACHE_NAME);
      cache.put(request, response.clone());
    }
    return response;
  } catch (_) {
    if (esPost) {
      // Sin red en POST → error JSON directo, sin intentar caché
      return new Response(
        JSON.stringify({ ok: false, error: 'Sin conexión' }),
        { status: 503, headers: { 'Content-Type': 'application/json' } }
      );
    }
    const cached = await caches.match(request);
    if (cached) return cached;
    return new Response(
      JSON.stringify({ ok: false, error: 'Sin conexión' }),
      { status: 503, headers: { 'Content-Type': 'application/json' } }
    );
  }
}
