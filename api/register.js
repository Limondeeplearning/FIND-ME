// POST /api/register — a participant registers name + seat (+ push subscription if allowed).
// mode: "register" (default) | "claim" ("This is my seat — get staff help") | "unclaim" (cancel that)
import { send, readJson, cleanSeat, SEAT_RE, DEVICE_RE, postAppsScript, methodGuard } from './_lib.js';

function cleanSubscription(sub) {
  if (!sub || typeof sub !== 'object') return '';
  const { endpoint, keys } = sub;
  if (typeof endpoint !== 'string' || !endpoint.startsWith('https://')) return '';
  if (!keys || typeof keys.p256dh !== 'string' || typeof keys.auth !== 'string') return '';
  const json = JSON.stringify({ endpoint, expirationTime: sub.expirationTime ?? null, keys: { p256dh: keys.p256dh, auth: keys.auth } });
  return json.length <= 2000 ? json : '';
}

export default async function handler(req, res) {
  if (!methodGuard(req, res, 'POST')) return;
  let body;
  try {
    body = await readJson(req);
  } catch {
    return send(res, 400, { ok: false, code: 'BAD_REQUEST' });
  }

  const seat = cleanSeat(body.seat);
  const name = String(body.name ?? '').replace(/\s+/g, ' ').trim().slice(0, 80);
  const deviceId = String(body.deviceId ?? '');
  const mode = ['register', 'claim', 'unclaim'].includes(body.mode) ? body.mode : 'register';

  if (!SEAT_RE.test(seat)) return send(res, 400, { ok: false, code: 'SEAT_INVALID' });
  if (!DEVICE_RE.test(deviceId)) return send(res, 400, { ok: false, code: 'BAD_REQUEST' });
  if (mode !== 'unclaim' && name.length < 2) return send(res, 400, { ok: false, code: 'NAME_REQUIRED' });

  try {
    const result = await postAppsScript({
      action: mode,
      seat,
      name,
      deviceId,
      subscription: cleanSubscription(body.subscription),
    });
    return send(res, result.ok ? 200 : 409, result);
  } catch (err) {
    console.error('register failed', err);
    return send(res, 502, { ok: false, code: 'UPSTREAM', error: 'Could not reach the registration sheet.' });
  }
}
