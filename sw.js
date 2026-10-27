/* EYECON service worker: only shows notifications and brings the app forward when one is tapped (no caching) */
self.addEventListener('install',()=>self.skipWaiting());
self.addEventListener('activate',e=>e.waitUntil(self.clients.claim()));
self.addEventListener('notificationclick',e=>{e.notification.close();const d=e.notification.data||{};
  e.waitUntil(self.clients.matchAll({type:'window',includeUncontrolled:true}).then(cs=>{for(const c of cs){c.postMessage({nt:d});if('focus' in c)return c.focus()}return self.clients.openWindow('./')}))});
