// JOURNALのお知らせ係。届いた合図をもとに、中身を取りに行って表示する
importScripts("config.js");
const GAS = self.JOURNAL_GAS;
self.addEventListener("install", (e) => self.skipWaiting());
self.addEventListener("activate", (e) => e.waitUntil(self.clients.claim()));
self.addEventListener("push", (event) => {
  event.waitUntil((async () => {
    let msg = {title: "JOURNAL", body: "お知らせがあります", url: "./"};
    try {
      const sub = await self.registration.pushManager.getSubscription();
      const r = await fetch(GAS + "?action=push_latest&endpoint=" + encodeURIComponent(sub ? sub.endpoint : ""));
      const j = await r.json();
      if (j && j.title) msg = j;
    } catch (e) {}
    await self.registration.showNotification(msg.title, {body: msg.body || "", icon: "icon-192.png", badge: "icon-192.png", data: {url: msg.url || "./"}});
  })());
});
self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  const url = (event.notification.data && event.notification.data.url) || "./";
  event.waitUntil((async () => {
    const list = await self.clients.matchAll({type: "window", includeUncontrolled: true});
    for (const c of list){ if ("focus" in c){ await c.navigate(url).catch(() => {}); return c.focus(); } }
    return self.clients.openWindow(url);
  })());
});
