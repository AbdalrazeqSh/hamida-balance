// خدمة عامل حقيقية: هذه الصفحة مستضافة خارج بيئة Claude (على GitHub Pages)،
// فتسجيل Service Worker هنا يعمل بشكل طبيعي ويتيح فتح التطبيق بدون اتصال من الصفر.
var CACHE_NAME = 'hamida-shell-v8';
var SHELL_FILES = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', function (event) {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      return cache.addAll(SHELL_FILES);
    })
  );
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (names) {
      return Promise.all(
        names.filter(function (n) { return n !== CACHE_NAME; }).map(function (n) { return caches.delete(n); })
      );
    }).then(function () { return self.clients.claim(); })
  );
});

// نفس الأصل (صفحة التطبيق وملفاته): شبكة أولاً مع تحديث التخزين المؤقت، ورجوع للتخزين عند الفشل.
// أصل آخر (مثل خطوط Google): نحاول الشبكة، ونخزّن ما ينجح، ونستخدم النسخة المخزنة إن وجدت عند الفشل،
// وإن لم توجد نتجاهل الفشل بصمت (الخط البديل في التصميم يكفي).
self.addEventListener('fetch', function (event) {
  if (event.request.method !== 'GET') return;
  var req = event.request;

  event.respondWith(
    fetch(req).then(function (res) {
      var copy = res.clone();
      caches.open(CACHE_NAME).then(function (cache) { cache.put(req, copy); });
      return res;
    }).catch(function () {
      return caches.match(req).then(function (cached) {
        if (cached) return cached;
        // طلب تنقّل (فتح الصفحة نفسها) بدون نسخة مطابقة: أرجع الصفحة الرئيسية المخزّنة
        if (req.mode === 'navigate') {
          return caches.match('./index.html');
        }
        return new Response('', { status: 504, statusText: 'offline' });
      });
    })
  );
});
