// ═══════════════════════════════════════════════════════════════
// Vale Digital — Service Worker v7.0
// Estrategia: cache-first para assets propios
// GAS: siempre red directa, sin intercepción del SW
// Novedad v7: escucha mensaje SKIP_WAITING para activación inmediata
// ═══════════════════════════════════════════════════════════════

const CACHE_NAME = 'vale-digital-v7.0';

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
    // NO llamar skipWaiting aquí — esperamos el mensaje explícito
    // para que el usuario controle cuándo se aplica la actualización
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

// ── MENSAJE: SKIP_WAITING ────────────────────────────────────────
// Recibido desde doActualizar() cuando el usuario toca "Actualizar"
self.addEventListener('message', event => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
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
