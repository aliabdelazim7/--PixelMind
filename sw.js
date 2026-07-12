const CACHE_NAME = 'leadflow-cache-v1';
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

// حدث جلب البيانات (Fetch Event): تقديم الملفات من الكاش عند انقطاع الشبكة
self.addEventListener('fetch', e => {
  // تخطي الطلبات التي ليست من نوع GET أو طلبات جوجل شيت لمنع التعارض مع الحفظ والمسح
  if (e.request.method !== 'GET' || e.request.url.includes('script.google.com')) {
    return;
  }

  e.respondWith(
    caches.match(e.request).then(cachedResponse => {
      if (cachedResponse) {
        return cachedResponse;
      }
      return fetch(e.request).then(networkResponse => {
        // التحقق من صلاحية الاستجابة قبل إضافتها للكاش
        if (!networkResponse || networkResponse.status !== 200 || networkResponse.type !== 'basic') {
          return networkResponse;
        }
        const responseToCache = networkResponse.clone();
        caches.open(CACHE_NAME).then(cache => {
          cache.put(e.request, responseToCache);
        });
        return networkResponse;
      }).catch(() => {
        // في حال انقطاع الشبكة تماماً
        return caches.match('./index.html');
      });
    })
  );
});
