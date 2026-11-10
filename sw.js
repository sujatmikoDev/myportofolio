// Service worker portofolio SujatmikoDev
// Strategi: ambil dari jaringan dulu (konten selalu baru), jatuh ke cache kalau offline.
// Naikkan angka VERSI setiap kali kamu mengganti index.html supaya cache lama dibersihkan.
const VERSI = 'sdev-v1';
const SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];
const HOST_BOLEH = ['cdn.tailwindcss.com', 'cdnjs.cloudflare.com', 'fonts.googleapis.com', 'fonts.gstatic.com'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSI).then(c => c.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSI).map(k => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;                       // POST ke Apps Script tidak disentuh
  const url = new URL(req.url);
  const sama = url.origin === self.location.origin;
  if (!sama && !HOST_BOLEH.includes(url.hostname)) return; // gambar Drive, YouTube, dll. lewat jaringan biasa
  e.respondWith(
    fetch(req)
      .then(res => {
        if (res && (res.ok || res.type === 'opaque')) {
          const salinan = res.clone();
          caches.open(VERSI).then(c => c.put(req, salinan));
        }
        return res;
      })
      .catch(() => caches.match(req).then(hit => hit || (req.mode === 'navigate' ? caches.match('./index.html') : undefined)))
  );
});
