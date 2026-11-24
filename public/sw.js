/* Find Me — service worker
   - Shows the lock-screen notification when the sheet sends an alert (web push) — always first
   - Tells any open Find Me page to start the full-screen flash
   - Confirms receipt to the sheet ("Alert via: Push ✓")
   - Keeps a small offline copy of the app so it still opens on weak venue Wi-Fi */
'use strict';

const CACHE = 'findme-v7';
const STATE = 'findme-state'; // small flags shared with the page (never deleted on update)
const SHELL = ['/', '/index.html', '/styles.css', '/app.js', '/manifest.webmanifest', '/icons/icon-192.png', '/icons/badge-96.png'];
const DEFAULT_MSG = 'Please proceed to the holding area on the right. Please find the staff to assist you.';

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE && k !== STATE).map((k) => caches.delete(k))))
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

// Remember that a push reached this phone (the page reads this to show "Confirmed"),
// and tell the sheet, so staff can see which phones are reachable.
async function recordReceipt(data) {
  try {
    const c = await caches.open(STATE);
    await c.put('/__pushok', new Response(JSON.stringify({ at: Date.now(), seat: data.seat || '' }), { headers: { 'Content-Type': 'application/json' } }));
  } catch { /* ignore */ }
  if (data.seat && data.device) {
    try {
      await fetch('/api/register', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ mode: 'pushok', seat: data.seat, deviceId: data.device }),
      });
    } catch { /* offline — the next push will report again */ }
  }
}

function tellPages(msg) {
  return self.clients.matchAll({ type: 'window', includeUncontrolled: true })
    .then((wins) => wins.forEach((c) => { try { c.postMessage(msg); } catch { /* page frozen */ } }))
    .catch(() => {});
}

self.addEventListener('push', (event) => {
  let data = {};
  try { data = event.data ? event.data.json() : {}; } catch { data = { message: event.data ? event.data.text() : '' }; }
  const alertId = String(data.alertId || Date.now());
  const type = data.type || 'alert';
  const title = data.title || "It's your turn";
  const body = data.message || DEFAULT_MSG;
  const isAlert = type === 'alert';
  const notice = { type: 'findme-notice', kind: type, seat: data.seat, serverTime: data.serverTime || 0 };

  event.waitUntil((async () => {
    // 1) ALWAYS show the notification first — nothing else may stop it.
    try {
      await self.registration.showNotification(title, {
        body,
        tag: isAlert ? 'findme-alert' : 'findme-info',
        renotify: true,
        requireInteraction: isAlert || (type !== 'granted' && type !== 'welcome'),
        silent: false,
        vibrate: isAlert ? [700, 200, 700, 200, 1200, 200, 1200] : (type === 'welcome' ? [200, 100, 200] : [400, 150, 400, 150, 400]),
        icon: '/icons/icon-192.png',
        badge: '/icons/badge-96.png',
        timestamp: Date.now(),
        data: isAlert ? { alertId, url: '/?alert=' + encodeURIComponent(alertId) } : { url: '/', notice },
      });
    } catch (err) { /* keep going: still update pages + receipt */ }

    // 2) Then let any open Find Me page react (flash / notice), and record the receipt.
    await tellPages(isAlert ? { type: 'findme-alert', alertId, seat: data.seat, title, message: body } : notice);
    await recordReceipt(data);
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
