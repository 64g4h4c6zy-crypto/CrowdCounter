// 客流统计 PWA Service Worker
var CACHE = 'crowd-counter-v2';

self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE).then(function(cache) {
      return cache.addAll([
        '/CrowdCounter/',
        '/CrowdCounter/index.html',
        '/CrowdCounter/index-ai.html',
        '/CrowdCounter/manifest.json'
      ]);
    })
  );
});

self.addEventListener('fetch', function(e) {
  e.respondWith(
    caches.match(e.request).then(function(r) {
      // 网络优先，失败才用缓存
      return fetch(e.request).then(function(response) {
        return response;
      }).catch(function() {
        return r;
      });
    })
  );
});
