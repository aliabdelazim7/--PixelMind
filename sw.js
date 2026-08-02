const CACHE_NAME = 'leadflow-cache-v3';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  'https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap',
  'https://cdn.jsdelivr.net/npm/chart.js'
];

// حدث التثبيت (Install Event): تخزين الملفات الثابتة في الكاش
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(ASSETS);
    }).then(() => self.skipWaiting())
  );
});

// حدث التفعيل (Activate Event): حذف الكاش القديم
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys.map(key => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// حدث جلب البيانات (Fetch Event): استراتيجية الشبكة أولاً (Network-First) لتفادي تجمد الكاش
self.addEventListener('fetch', e => {
  // تخطي الطلبات التي ليست من نوع GET أو طلبات جوجل شيت لمنع التعارض مع قاعدة البيانات
  if (e.request.method !== 'GET' || e.request.url.includes('script.google.com')) {
    return;
  }

  e.respondWith(
    fetch(e.request).then(networkResponse => {
      // إذا كانت الاستجابة سليمة، نحدث الكاش بالنسخة الجديدة ونرجعها
      if (networkResponse && networkResponse.status === 200) {
        const responseToCache = networkResponse.clone();
        caches.open(CACHE_NAME).then(cache => {
          cache.put(e.request, responseToCache);
        });
      }
      return networkResponse;
    }).catch(() => {
      // في حال انقطاع الإنترنت تماماً، نقرأ من الكاش المتاح
      return caches.match(e.request).then(cachedResponse => {
        if (cachedResponse) {
          return cachedResponse;
        }
        // كحل نهائي إذا لم نجد الملف بالكاش
        return caches.match('./index.html');
      });
    })
  );
});
