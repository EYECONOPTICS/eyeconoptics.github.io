/* EYECON service worker (v136): keeps a copy of the app on the device so it opens without internet,
   shows notifications and brings the app forward when one is tapped.
   Pages and config: network first (fresh when online, the saved copy when offline or after 5s).
   Icons, images, fonts and libraries: the saved copy first, refreshed in the background.
   Calls to the Google Apps Script server are never touched (the app's own outbox handles them). */
var V = 'eyecon-shell-v1';
var CORE = ['./', 'index.html', 'config.js', 'manifest.webmanifest', 'favicon.ico',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/apple-touch-icon.png', 'icons/favicon-32.png', 'icons/favicon-16.png',
  'https://fonts.googleapis.com/css2?family=Noto+Sans:wght@400;500;600;700&family=Noto+Sans+Arabic:wght@400;500;600;700&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/qrcode-generator/1.4.4/qrcode.min.js'];
var STATIC_HOSTS = /^https:\/\/(fonts\.googleapis\.com|fonts\.gstatic\.com|cdnjs\.cloudflare\.com)\//;

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(V).then(function (c) {
    return Promise.all(CORE.map(function (u) { return c.add(new Request(u, { cache: 'reload' })).catch(function () {}); }));
  }).then(function () { return self.skipWaiting(); }));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) {
    return Promise.all(ks.filter(function (k) { return k.indexOf('eyecon-shell-') === 0 && k !== V; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

function put(req, res) {
  if (res && (res.ok || res.type === 'opaque')) { var cp = res.clone(); caches.open(V).then(function (c) { c.put(req, cp); }).catch(function () {}); }
  return res;
}
function fromCache(req, nav) {
  return caches.match(req, { ignoreSearch: true }).then(function (r) {
    if (r || !nav) return r;
    return caches.match('index.html', { ignoreSearch: true }).then(function (x) { return x || caches.match('./'); });
  });
}
function networkFirst(req, nav) {
  return new Promise(function (resolve) {
    var done = false, fin = function (r) { if (!done && r) { done = true; resolve(r); } };
    var t = setTimeout(function () { fromCache(req, nav).then(fin); }, 5000);
    fetch(req).then(function (res) { clearTimeout(t); put(req, res); if (done) return; if (res.ok) fin(res); else fromCache(req, nav).then(function (c) { done ? 0 : fin(c || res); }); },
      function () { clearTimeout(t); fromCache(req, nav).then(function (c) { if (!done) { done = true; resolve(c || Response.error()); } }); });
  });
}
function cacheFirst(req) {
  return caches.match(req).then(function (c) {
    var net = fetch(req).then(function (res) { return put(req, res); }).catch(function () { return c || Response.error(); });
    return c || net;
  });
}

self.addEventListener('fetch', function (e) {
  var req = e.request; if (req.method !== 'GET') return;
  var u = req.url;
  if (STATIC_HOSTS.test(u)) { e.respondWith(cacheFirst(req)); return; }
  if (u.indexOf(self.registration.scope) !== 0) return;              // other hosts (Apps Script, Google Drive…) go straight to the network
  var path = u.slice(self.registration.scope.length).split(/[?#]/)[0];
  if (req.mode === 'navigate' || path === '' || /\.(html|js|webmanifest)$/.test(path)) { e.respondWith(networkFirst(req, req.mode === 'navigate')); return; }
  e.respondWith(cacheFirst(req));
});

self.addEventListener('notificationclick', function (e) {
  e.notification.close(); var d = e.notification.data || {};
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (cs) {
    for (var i = 0; i < cs.length; i++) { cs[i].postMessage({ nt: d }); if ('focus' in cs[i]) return cs[i].focus(); }
    return self.clients.openWindow('./');
  }));
});
