// POST /api/push — called by the Google Sheet (Apps Script) to send push notifications.
// Body: { secret, title, message, targets: [{ type, seat, alertId, device, subscription, title?, message?, serverTime }] }
// type: alert (your turn) | unregistered (register again) | granted (seat is yours) | dismissed (check your seat) | welcome
import { env, send, readJson, safeEqual, methodGuard } from './_lib.js';
import { initPush, sendPush } from './_push.js';

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
  if (!initPush()) return send(res, 500, { ok: false, error: 'VAPID keys are not set in Vercel' });

  const targets = Array.isArray(body.targets) ? body.targets.slice(0, 500) : [];
  const title = String(body.title || "It's your turn");
  const message = String(body.message || '');

  const results = await Promise.allSettled(
    targets.map((t) =>
      sendPush(t.subscription, {
        type: t.type,
        seat: t.seat,
        alertId: t.alertId,
        device: t.device,
        title: t.title || title,
        message: t.message || message,
        serverTime: t.serverTime,
      })
    )
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
