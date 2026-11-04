// Offline-Fähigkeit: App-Dateien zwischenspeichern. Die Excel-Daten liegen nicht hier,
// sondern in der IndexedDB des Geräts.
const CACHE = 'unser-geld-v0.12.0';
const DATEIEN = ['./', 'index.html', 'app.css', 'app.js', 'kern/xlsx.js', 'kern/lesen.js', 'kern/rechnen.js', 'kern/verstehen.js', 'anmeldung.js', 'onedrive.js', 'config.js', 'pc.js', 'kern/csv.js', 'kern/privat.js', 'meins.js', 'kern/kai.js', 'kai.js', 'hilfe.js', 'kai.webmanifest',
  'manifest.webmanifest', 'icons/icon-192.png', 'icons/apple-touch-icon.png'];
self.addEventListener('install', (e) => e.waitUntil(caches.open(CACHE).then((c) => c.addAll(DATEIEN)).then(() => self.skipWaiting())));
self.addEventListener('activate', (e) => e.waitUntil(caches.keys()
  .then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim())));
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== location.origin) return;
  // Netz zuerst (damit Updates sofort ankommen), sonst aus dem Zwischenspeicher
  e.respondWith(fetch(e.request, { cache: 'no-cache' }).then((r) => { const k = r.clone(); caches.open(CACHE).then((c) => c.put(e.request, k)); return r; })
    .catch(() => caches.match(e.request)));
});
