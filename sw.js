/* 앱 화면(껍데기)만 보관한다. 시트 데이터(script.google.com)는 절대 가로채거나 보관하지 않는다.
   화면은 «네트워크 우선» 이라 새로 발행하면 바로 반영되고, 연결이 없을 때만 보관본을 보여준다. */
const CACHE = 'pw-shell-v1';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon.svg', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).catch(() => {}).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request, url = new URL(req.url);
  if(req.method !== 'GET' || url.origin !== self.location.origin) return;   // 외부 요청(시트·폰트)은 그대로 통과
  e.respondWith(
    fetch(req).then(res => {
      if(res && res.ok){ const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {}); }
      return res;
    }).catch(() => caches.match(req).then(hit => hit || (req.mode === 'navigate' ? caches.match('./index.html') : undefined)))
  );
});
