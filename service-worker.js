/* บัญชีเพจครูพร้อมสอน — Service Worker */
const CACHE = 'krupromson-acct-v2';
const SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).catch(()=>{}));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
  ));
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  // ข้อมูลจาก API (Google) — ให้วิ่งเน็ตตรงๆ ไม่แคช
  if (url.hostname.indexOf('google.com') >= 0 || url.hostname.indexOf('googleusercontent.com') >= 0) return;
  // ฟอนต์/ไฟล์เปลือกแอป — แคชก่อน ค่อยเน็ต
  e.respondWith(
    caches.match(e.request).then(hit => hit || fetch(e.request).then(res => {
      if (e.request.method === 'GET' && res && res.status === 200 && url.origin === location.origin) {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, copy));
      }
      return res;
    }).catch(() => caches.match('./index.html')))
  );
});
