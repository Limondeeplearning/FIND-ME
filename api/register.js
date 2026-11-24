// POST /api/register — everything a participant's phone sends to the sheet.
// mode:
//   "register" (default) name + seat (+ push subscription) → then a "You're all set" test push
//   "claim"    "This is my seat — get staff help"
//   "unclaim"  cancel that claim
//   "mobile"   save / remove a mobile number for SMS alerts
//   "pushok"   sent by the service worker when a push arrives → phone confirmed reachable
import { send, readJson, cleanSeat, SEAT_RE, DEVICE_RE, postAppsScript, methodGuard } from './_lib.js';

const MODES = ['register', 'claim', 'unclaim', 'mobile', 'pushok'];

function cleanSubscription(sub) {
  if (!sub || typeof sub !== 'object') return '';
  const { endpoint, keys } = sub;
  if (typeof endpoint !== 'string' || !endpoint.startsWith('https://')) return '';
  if (!keys || typeof keys.p256dh !== 'string' || typeof keys.auth !== 'string') return '';
  const json = JSON.stringify({ endpoint, expirationTime: sub.expirationTime ?? null, keys: { p256dh: keys.p256dh, auth: keys.auth } });
  return json.length <= 2000 ? json : '';
}

/** Sends the automatic "You're all set" push so the phone can confirm notifications really arrive. */
async function sendWelcome(subscription, seat, deviceId, serverTime) {
  try {
    const { sendPush } = await import('./_push.js');
    await sendPush(subscription, {
      type: 'welcome',
      seat,
      device: deviceId,
      alertId: 'welcome-' + Date.now(),
      title: "✅ You're all set",
      message: "Find Me will alert you here when it's your turn. Keep notifications on.",
      serverTime,
    });
    return true;
  } catch (err) {
    console.error('welcome push failed', err?.statusCode, err?.body || err?.message);
    return false;
  }
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
  const mode = MODES.includes(body.mode) ? body.mode : 'register';

  if (!SEAT_RE.test(seat)) return send(res, 400, { ok: false, code: 'SEAT_INVALID' });
  if (!DEVICE_RE.test(deviceId)) return send(res, 400, { ok: false, code: 'BAD_REQUEST' });
  if ((mode === 'register' || mode === 'claim') && name.length < 2) return send(res, 400, { ok: false, code: 'NAME_REQUIRED' });

  const payload = { action: mode, seat, deviceId };
  if (mode === 'register' || mode === 'claim') {
    payload.name = name;
    payload.subscription = cleanSubscription(body.subscription);
  }
  if (mode === 'mobile') payload.mobile = body.mobile === '' || body.mobile === null ? '' : String(body.mobile ?? '').slice(0, 20);

  try {
    const result = await postAppsScript(payload);
    const nowRegistered = result.ok && (mode === 'register' || (mode === 'claim' && result.registered));
    if (nowRegistered && payload.subscription) {
      result.welcomeSent = await sendWelcome(payload.subscription, seat, deviceId, result.serverTime);
    }
    return send(res, result.ok ? 200 : 409, result);
  } catch (err) {
    console.error(mode + ' failed', err);
    return send(res, 502, { ok: false, code: 'UPSTREAM', error: 'Could not reach the registration sheet.' });
  }
}
