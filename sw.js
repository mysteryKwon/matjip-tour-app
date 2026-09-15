/* 맛집투어관리 앱 — 서비스워커 */
var APP_VERSION = '1.0.0';
var CACHE_NAME = 'mtapp-cache-v' + APP_VERSION;

// config.json은 캐시하지 않는다 (API URL을 항상 최신으로 가져오기 위함)
var PRECACHE = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

self.addEventListener('install', function(event){
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache){ return cache.addAll(PRECACHE); })
  );
});

self.addEventListener('activate', function(event){
  event.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.filter(function(k){ return k !== CACHE_NAME; }).map(function(k){ return caches.delete(k); }));
    }).then(function(){ return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function(event){
  var url = event.request.url;
  // config.json과 Apps Script API 호출은 항상 네트워크에서 최신으로 가져온다
  if(url.indexOf('config.json') > -1 || url.indexOf('script.google.com') > -1){
    event.respondWith(fetch(event.request).catch(function(){ return caches.match(event.request); }));
    return;
  }
  event.respondWith(
    caches.match(event.request).then(function(cached){
      if(cached) return cached;
      return fetch(event.request).then(function(res){
        if(event.request.method === 'GET' && res && res.status === 200 && res.type === 'basic'){
          var resClone = res.clone();
          caches.open(CACHE_NAME).then(function(cache){ cache.put(event.request, resClone); });
        }
        return res;
      }).catch(function(){
        if(event.request.mode === 'navigate') return caches.match('./index.html');
      });
    })
  );
});
