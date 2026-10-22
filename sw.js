// PO service worker — בכל עדכון של האפליקציה מעלים את מספר הגרסה
const VERSION = 'po-v4';
const MODELS = 'po-models-v1'; // מנוע הזיהוי (cdn.jsdelivr) — נשמר כדי שהאפליקציה תעלה מהר וגם בלי רשת
const APP_FILES = ['./', 'index.html', 'manifest.webmanifest', 'logo.png',
  'icon-180.png', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(APP_FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== VERSION && k !== MODELS).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // מנוע הזיהוי והמודלים: קודם מה-cache, ואם אין — מהרשת ושומרים
  if (url.hostname === 'cdn.jsdelivr.net' && url.pathname.includes('@vladmandic/face-api')) {
    e.respondWith(caches.open(MODELS).then(c => c.match(req).then(hit => hit ||
      fetch(req).then(res => { if (res.ok) c.put(req, res.clone()); return res; }))));
    return;
  }

  // שאר הבקשות החיצוניות — ישר לרשת
  if (url.origin !== self.location.origin) return;

  // קבצי האפליקציה: קודם רשת, ואם אין חיבור — מה-cache
  e.respondWith(fetch(req).then(res => {
    if (res.ok) { const copy = res.clone(); caches.open(VERSION).then(c => c.put(req, copy)); }
    return res;
  }).catch(() => caches.match(req).then(hit => hit || caches.match('index.html'))));
});
