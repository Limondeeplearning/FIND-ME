// POST /api/push — called by the Google Sheet (Apps Script) to send push notifications.
// Body: { secret, title, message, targets: [{ type, seat, alertId, subscription, title?, message?, serverTime }] }
// type: alert (your turn) | unregistered (register again) | granted (seat is yours) | dismissed (check your seat)
import webpush from 'web-push';
import { env, send, readJson, safeEqual, methodGuard } from './_lib.js';

export const config = { maxDuration: 30 };

export default async function handler(req, res) {
  if (!methodGuard(req, res, 'POST')) return;

  let body;
  try {
    body = await readJson(req, 1024 * 1024);
  } catch {
    return send(res, 400, { ok: false, error: 'Bad JSON' });
  }
  if (!safeEqual(body.secret, env('SHARED_SECRET'))) return send(res, 401, { ok: false, error: 'Unauthorized' });

  const publicKey = env('VAPID_PUBLIC_KEY');
  const privateKey = env('VAPID_PRIVATE_KEY');
  if (!publicKey || !privateKey) return send(res, 500, { ok: false, error: 'VAPID keys are not set in Vercel' });
  webpush.setVapidDetails(env('VAPID_SUBJECT', 'mailto:admin@example.com'), publicKey, privateKey);

  const targets = Array.isArray(body.targets) ? body.targets.slice(0, 500) : [];
  const title = String(body.title || "It's your turn").slice(0, 80);
  const message = String(body.message || '').slice(0, 300);

  const results = await Promise.allSettled(
    targets.map(async (t) => {
      const sub = typeof t.subscription === 'string' ? JSON.parse(t.subscription) : t.subscription;
      const payload = JSON.stringify({
        type: ['alert', 'unregistered', 'granted', 'dismissed'].includes(t.type) ? t.type : 'alert',
        seat: t.seat,
        alertId: t.alertId,
        title: String(t.title || title).slice(0, 80),
        message: String(t.message || message).slice(0, 300),
        serverTime: Number(t.serverTime) || Date.now(),
      });
      await webpush.sendNotification(sub, payload, { TTL: 600, urgency: 'high', topic: 'findme' });
      return t.seat;
    })
  );

  const sent = [];
  const expired = [];
  const expiredIdx = [];
  const failed = [];
  results.forEach((r, i) => {
    const seat = targets[i]?.seat;
    if (r.status === 'fulfilled') sent.push(seat);
    else if ([404, 410].includes(r.reason?.statusCode)) { expired.push(seat); expiredIdx.push(i); }
    else {
      failed.push(seat);
      console.error('push failed', seat, r.reason?.statusCode, r.reason?.body || r.reason?.message);
    }
  });

  return send(res, 200, { ok: true, sent, expired, expiredIdx, failed });
}
