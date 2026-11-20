/* Tiny IndexedDB wrapper shared by the page and the service worker.
   Classic script (no modules) so the service worker can importScripts() it. */
(function (root) {
  const DB_NAME = 'minutes';
  const DB_VERSION = 1;
  let dbPromise = null;

  function open() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains('kv')) db.createObjectStore('kv');
        if (!db.objectStoreNames.contains('jobs')) db.createObjectStore('jobs', { keyPath: 'id' });
        if (!db.objectStoreNames.contains('audio')) db.createObjectStore('audio');
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    return dbPromise;
  }

  function reqP(r) {
    return new Promise((resolve, reject) => {
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    });
  }

  async function get(store, key) {
    const db = await open();
    return reqP(db.transaction(store).objectStore(store).get(key));
  }
  async function put(store, value, key) {
    const db = await open();
    const t = db.transaction(store, 'readwrite');
    const done = new Promise((res, rej) => { t.oncomplete = res; t.onerror = () => rej(t.error); });
    key === undefined ? t.objectStore(store).put(value) : t.objectStore(store).put(value, key);
    return done;
  }
  async function del(store, key) {
    const db = await open();
    const t = db.transaction(store, 'readwrite');
    const done = new Promise((res, rej) => { t.oncomplete = res; t.onerror = () => rej(t.error); });
    t.objectStore(store).delete(key);
    return done;
  }
  async function all(store) {
    const db = await open();
    return reqP(db.transaction(store).objectStore(store).getAll());
  }

  root.MinutesDB = { open, get, put, del, all };
})(typeof self !== 'undefined' ? self : window);
