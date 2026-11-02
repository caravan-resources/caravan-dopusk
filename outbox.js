/* Очередь отправки без связи (IndexedDB). Общая для страниц (checklist.html, bvr-form.html)
   и service worker (sw.js): запись, которую не удалось отправить, лежит на телефоне и уходит
   на сервер, как только появится связь. Повторная отправка безопасна: у каждой записи есть
   clientId, сервер второй раз её не запишет.
   item = { id, createdAt, body, label, after } — body это то, что уходит на сервер как есть. */
(function (g) {
  const DB = "operbaza-outbox", STORE = "items";

  function open() {
    return new Promise((res, rej) => {
      const r = indexedDB.open(DB, 1);
      r.onupgradeneeded = () => { r.result.createObjectStore(STORE, { keyPath: "id" }); };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
  }
  function run(mode, fn) {
    return open().then(db => new Promise((res, rej) => {
      const t = db.transaction(STORE, mode);
      let out;
      const q = fn(t.objectStore(STORE));
      if (q) q.onsuccess = () => { out = q.result; };
      t.oncomplete = () => { db.close(); res(out); };
      t.onerror = t.onabort = () => { db.close(); rej(t.error); };
    }));
  }

  const OB = {
    add: item => run("readwrite", s => s.put(item)),
    all: () => run("readonly", s => s.getAll()).then(a => (a || []).sort((x, y) => x.createdAt - y.createdAt)),
    remove: id => run("readwrite", s => s.delete(id)),
    // post(body) -> Promise<json>; при отсутствии связи обязан бросить исключение.
    // Нет связи: останавливаемся, остальное уйдёт позже. Сервер отклонил запись: помечаем попытку и идём дальше.
    flush: async function (post, onSent) {
      const items = await OB.all();
      let sent = 0, rejected = 0;
      for (const it of items) {
        let r;
        try { r = await post(it.body); }
        catch (e) { return { sent: sent, left: items.length - sent, offline: true, rejected: rejected }; }
        if (r && r.ok) {
          await OB.remove(it.id); sent++;
          if (onSent) { try { await onSent(it, r); } catch (e) { /* вторичное действие */ } }
        } else {
          it.attempts = (it.attempts || 0) + 1;
          it.lastError = String((r && r.error) || "ошибка");
          await OB.add(it); rejected++;
        }
      }
      return { sent: sent, left: items.length - sent, offline: false, rejected: rejected };
    },
  };
  g.OB = OB;
})(typeof self !== "undefined" ? self : window);
