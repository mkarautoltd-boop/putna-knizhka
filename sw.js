/* минимален service worker – само за да може приложението да се инсталира на компютър; нищо не пази в кеша */
self.addEventListener('install',()=>self.skipWaiting());
self.addEventListener('activate',e=>e.waitUntil(self.clients.claim()));
self.addEventListener('fetch',e=>{const r=e.request;if(r.method==='GET'&&new URL(r.url).origin===location.origin)e.respondWith(fetch(r))});
