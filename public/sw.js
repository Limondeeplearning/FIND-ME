/* Find Me — service worker
   - Shows the lock-screen notification when the sheet sends an alert (web push)
   - Tells any open Find Me page to start the full-screen flash
   - Keeps a small offline copy of the app so it still opens on weak venue Wi-Fi */
'use strict';

const CACHE = 'findme-v2';
const SHELL = ['/', '/index.html', '/styles.css', '/app.js', '/manifest.webmanifest', '/icons/icon-192.png', '/icons/badge-96.png'];
const DEFAULT_MSG = 'Please proceed to the holding area on the right. Please find the staff to assist you.';

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Network first (so updates show), cached copy if the network is down. API calls are never cached.
self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin || url.pathname.startsWith('/api/')) return;
  event.respondWith(
    fetch(req)
      .then((res) => {
        if (res.ok && SHELL.includes(url.pathname)) {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(req, copy));
        }
        return res;
      })
      .catch(() => caches.match(req).then((hit) => hit || caches.match('/index.html')))
  );
});

self.addEventListener('push', (event) => {
  let data = {};
  try { data = event.data ? event.data.json() : {}; } catch { data = { message: event.data ? event.data.text() : '' }; }
  const alertId = String(data.alertId || Date.now());
  const type = data.type || 'alert';
  const title = data.title || "It's your turn";
  const body = data.message || DEFAULT_MSG;

  // Staff notices: register again / seat is yours / check your seat.
  if (type !== 'alert') {
    const notice = { type: 'findme-notice', kind: type, seat: data.seat, serverTime: data.serverTime || 0 };
    event.waitUntil((async () => {
      const wins = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
      wins.forEach((c) => c.postMessage(notice));
      await self.registration.showNotification(title, {
        body,
        tag: 'findme-info',
        renotify: true,
        requireInteraction: type !== 'granted',
        vibrate: [400, 150, 400, 150, 400],
        icon: '/icons/icon-192.png',
        badge: '/icons/badge-96.png',
        timestamp: Date.now(),
        data: { url: '/', notice },
      });
    })());
    return;
  }

  event.waitUntil((async () => {
    // Open pages start flashing immediately.
    const wins = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    wins.forEach((c) => c.postMessage({ type: 'findme-alert', alertId, seat: data.seat, title, message: body }));
    // Always show a notification (required by iOS/Chrome for every push).
    await self.registration.showNotification(title, {
      body,
      tag: 'findme-alert',
      renotify: true,
      requireInteraction: true,
      vibrate: [700, 200, 700, 200, 1200, 200, 1200],
      icon: '/icons/icon-192.png',
      badge: '/icons/badge-96.png',
      timestamp: Date.now(),
      data: { alertId, url: '/?alert=' + encodeURIComponent(alertId) },
    });
  })());
});

self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const d = event.notification.data || {};
  const tag = event.notification.tag;
  if (tag !== 'findme-alert' && tag !== 'findme-info') return;
  event.waitUntil((async () => {
    const wins = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    for (const c of wins) {
      if ('focus' in c) {
        c.postMessage(tag === 'findme-alert' ? { type: 'findme-alert', alertId: d.alertId } : d.notice);
        return c.focus();
      }
    }
    return self.clients.openWindow(d.url || '/');
  })());
});
