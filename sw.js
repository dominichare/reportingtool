/* Minutes service worker: offline app shell + Android Share-sheet target. */
importScripts('js/db.js');

const VERSION = 'minutes-v2';
const SHELL = [
  './',
  'index.html',
  'manifest.webmanifest',
  'css/app.css',
  'fonts/inter.woff2',
  'js/db.js',
  'js/app.js',
  'js/settings.js',
  'js/providers.js',
  'js/note.js',
  'js/docx.js',
  'js/graph.js',
  'js/help.js',
  'vendor/anthropic-sdk-1.6.0.mjs',
  'icons/icon-192.png',
  'icons/icon-512.png',
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    for (const k of await caches.keys()) if (k !== VERSION) await caches.delete(k);
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  const req = event.request;
  const url = new URL(req.url);

  // A recording shared from Easy Recorder (or any app) via the Android Share sheet.
  if (req.method === 'POST' && url.origin === location.origin && url.pathname.endsWith('/share-target')) {
    event.respondWith((async () => {
      const base = new URL('./', self.registration.scope);
      try {
        const form = await req.formData();
        const files = form.getAll('audio').filter(f => f && typeof f === 'object' && f.size > 0);
        if (!files.length) return Response.redirect(new URL('./?shared=none', base).href, 303);
        await self.MinutesDB.put('kv', { file: files[0], at: Date.now() }, 'pendingShare');
        return Response.redirect(new URL('./?shared=1', base).href, 303);
      } catch (e) {
        return Response.redirect(new URL('./?shared=none', base).href, 303);
      }
    })());
    return;
  }

  if (req.method !== 'GET' || url.origin !== location.origin) return;

  // App shell: network first (so updates arrive), falling back to the cache when offline.
  event.respondWith((async () => {
    try {
      const res = await fetch(req);
      if (res.ok) {
        const c = await caches.open(VERSION);
        c.put(req, res.clone());
      }
      return res;
    } catch {
      const cached = await caches.match(req, { ignoreSearch: true });
      return cached || caches.match('./');
    }
  })());
});
