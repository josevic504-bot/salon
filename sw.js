// Eventos VM: permite instalar la app y abrirla sin conexión.
// La página se pide SIEMPRE a internet sin usar copias guardadas (así cada actualización de GitHub llega al abrir);
// solo si no hay señal se usa la copia guardada. Las llamadas a Apps Script nunca se guardan.
const CACHE = 'eventos-vm-v2';
const BASE = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(BASE.map(u => new Request(u, { cache: 'reload' })))).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin && !url.hostname.endsWith('fonts.googleapis.com') && !url.hostname.endsWith('fonts.gstatic.com')) return;
  const pagina = req.mode === 'navigate' || url.pathname.endsWith('/') || url.pathname.endsWith('.html');
  const pedir = pagina ? fetch(new Request(req.url, { cache: 'no-store', credentials: 'same-origin' })) : fetch(req);

  e.respondWith(
    pedir
      .then(res => {
        if (res && res.ok && !url.search) {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(req, copy));
        }
        return res;
      })
      .catch(() => caches.match(req, { ignoreSearch: true }).then(r => r || (pagina ? caches.match('./index.html') : undefined)))
  );
});
