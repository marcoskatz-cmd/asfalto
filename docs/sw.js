// SW minimo: hace la PWA instalable y sirve la cascara sin red. Los datos siempre van a la API.
const CACHE = 'asfalto-v1';
const ASSETS = ['./', './index.html', './manifest.json', './logo.png', './icon-192.png', './icon-512.png'];
// Estilo común de la App INGECO: red primero (siempre la última), caché solo sin señal.
const COMUN = 'https://app.ingecosa.com.ar/ingeco.';
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS.map(a => new Request(a, { cache: 'reload' })))).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method === 'GET' && e.request.url.startsWith(COMUN)) {
    e.respondWith(fetch(e.request).then(r => { if (r && r.ok) { const copia = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copia)); } return r; }).catch(() => caches.match(e.request)));
    return;
  }
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  e.respondWith(fetch(e.request, { cache: 'no-cache' }).then(r => { const c = r.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); return r; }).catch(() => caches.match(e.request)));
});
