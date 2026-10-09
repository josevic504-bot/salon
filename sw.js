/* Guarda el sistema en el teléfono para que abra al instante.
   Las consultas a Google (datos del salón) nunca se guardan aquí. */
// El nombre incluye la carpeta, para que dos instalaciones en la misma cuenta no se mezclen
const VERSION = 'salon-v3@' + self.registration.scope;
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(
  caches.keys().then(ks => Promise.all(ks.filter(k => k.endsWith('@' + self.registration.scope) && k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim())
));
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const u = new URL(r.url);
  if (/script\.google|googleusercontent/.test(u.hostname)) return;
  const externo = /cdn\.jsdelivr\.net|cdnjs\.cloudflare\.com|fonts\.googleapis\.com|fonts\.gstatic\.com/.test(u.hostname);
  if (externo) {
    // Librerías con versión fija: se descargan una sola vez
    e.respondWith(caches.open(VERSION).then(c => c.match(r).then(m => m || fetch(r).then(res => { if (res.ok || res.type === 'opaque') c.put(r, res.clone()); return res; }))));
  } else if (u.origin === location.origin) {
    // La página: se muestra la copia guardada y se actualiza en segundo plano
    e.respondWith(caches.open(VERSION).then(c => c.match(r).then(m => {
      const red = fetch(r).then(res => { if (res.ok) c.put(r, res.clone()); return res; }).catch(() => m);
      return m || red;
    })));
  }
});
