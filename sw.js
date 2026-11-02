/* Service worker Опербазы: страница открывается без связи + очередь отправляется сама.
   1) Страницы checklist.html, bvr-form.html и outbox.js кэшируются: после первого открытия
      со связью QR-ссылка открывается и в зоне без связи. Онлайн всегда берётся свежая версия
      (сеть в приоритете, при слабой связи через 4 с отдаётся копия из памяти).
   2) Background Sync (Chrome на Android): если запись осталась в очереди, а страницу закрыли,
      телефон сам отправит её, как только появится связь. На iPhone этого нет: там очередь
      уйдёт при следующем открытии страницы. */
importScripts("outbox.js");

const CACHE = "operbaza-offline-v1";
const FILES = ["checklist.html", "bvr-form.html", "outbox.js"];
const SHEET_URL = "https://script.google.com/macros/s/AKfycbwhdMV20k8US2BXUz1Xk2XCXE9MEwpHpQaadTF9xEfzbU1--dHL-Zuj06K2qXW-rFtKGw/exec";

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE)
    .then(c => Promise.all(FILES.map(f => c.add(new Request(f, { cache: "reload" })).catch(() => {}))))
    .then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener("fetch", e => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== self.location.origin) return;
  const name = u.pathname.split("/").pop();
  if (FILES.indexOf(name) < 0) return;
  const key = u.origin + u.pathname;
  e.respondWith((async () => {
    const cache = await caches.open(CACHE);
    try {
      const ctrl = new AbortController();
      const t = setTimeout(() => ctrl.abort(), 4000);
      const r = await fetch(e.request, { signal: ctrl.signal });
      clearTimeout(t);
      if (r.ok) cache.put(key, r.clone());
      return r;
    } catch (err) {
      const m = await cache.match(key);
      return m || Response.error();
    }
  })());
});

self.addEventListener("sync", e => {
  if (e.tag !== "operbaza-outbox") return;
  e.waitUntil(OB.flush(async body => {
    const r = await fetch(SHEET_URL, { method: "POST", body: JSON.stringify(body) });
    return r.json();
  }).then(res => { if (res.offline) throw new Error("offline"); }));   // исключение = браузер повторит позже
});
