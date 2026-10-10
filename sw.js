/* Guarda el sistema en el teléfono para que abra al instante.
   Las consultas a Google (datos del salón) nunca se guardan aquí. */
// El nombre incluye la carpeta, para que dos instalaciones en la misma cuenta no se mezclen
const VERSION = 'salon-v4@' + self.registration.scope;
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
  } else if (u.origin === location.origin && (r.mode === 'navigate' || r.destination === 'document')) {
    // La página: siempre la versión más reciente publicada (así los cambios llegan al instante).
    // La copia guardada se usa solo sin internet o si la red tarda más de 4 segundos.
    e.respondWith(caches.open(VERSION).then(c => {
      const red = fetch(u.href, { cache: 'no-cache', credentials: 'same-origin' }).then(res => { if (res.ok) c.put(r, res.clone()); return res; });
      return c.match(r).then(m => !m ? red : Promise.race([red.catch(() => m), new Promise(ok => setTimeout(() => ok(m), 4000))]));
    }));
  } else if (u.origin === location.origin) {
    // Íconos y manifiesto: copia guardada y se actualiza en segundo plano
    e.respondWith(caches.open(VERSION).then(c => c.match(r).then(m => {
      const red = fetch(r).then(res => { if (res.ok) c.put(r, res.clone()); return res; }).catch(() => m);
      return m || red;
    })));
  }
});
