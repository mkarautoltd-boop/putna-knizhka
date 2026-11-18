/* service worker: инсталиране на компютър и известия на телефона; нищо не пази в кеша */
self.addEventListener('install',()=>self.skipWaiting());
self.addEventListener('activate',e=>e.waitUntil(self.clients.claim()));
self.addEventListener('fetch',e=>{const r=e.request;if(r.method==='GET'&&new URL(r.url).origin===location.origin)e.respondWith(fetch(r))});
/* известие от сървъра (Firebase): {notification:{title,body,tag}, fcmOptions:{link}} */
self.addEventListener('push',e=>{
  let d={};try{d=e.data?e.data.json():{}}catch(x){d={notification:{title:'Пътна книжка',body:e.data?e.data.text():''}}}
  const n=d.notification||{},link=(d.fcmOptions&&d.fcmOptions.link)||(d.data&&d.data.link)||'putna.html';
  e.waitUntil(self.registration.showNotification(n.title||'Пътна книжка',{body:n.body||'',icon:'icon-192.png',badge:'icon-192.png',tag:n.tag||undefined,data:{link}}));
});
self.addEventListener('notificationclick',e=>{
  e.notification.close();const link=(e.notification.data&&e.notification.data.link)||'putna.html';
  e.waitUntil(self.clients.matchAll({type:'window',includeUncontrolled:true}).then(l=>{for(const c of l){if('focus' in c)return c.focus()}return self.clients.openWindow(link)}));
});
