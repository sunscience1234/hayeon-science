// 하연통합과학학원 학습도우미 Service Worker v3
var CACHE_NAME = 'hayeon-v3';

self.addEventListener('install', function(e) {
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.map(function(key){ return caches.delete(key); }));
    }).then(function(){ return self.clients.claim(); })
  );
});

// 항상 네트워크에서 최신 버전 가져오기
self.addEventListener('fetch', function(e) {
  e.respondWith(fetch(e.request).catch(function(){
    return caches.match(e.request);
  }));
});

self.addEventListener('push', function(e) {
  var data = {};
  try { data = e.data.json(); } catch(err) {
    data = { title: '하연통합과학학원', body: e.data ? e.data.text() : '새 알림!' };
  }
  e.waitUntil(self.registration.showNotification(data.title || '⚗️ 하연통합과학학원', {
    body: data.body || '새 알림이 있어요!',
    icon: '/icon-192.png',
    badge: '/icon-192.png',
    vibrate: [200, 100, 200],
    data: { url: data.url || '/' }
  }));
});

self.addEventListener('notificationclick', function(e) {
  e.notification.close();
  e.waitUntil(clients.openWindow('/'));
});
