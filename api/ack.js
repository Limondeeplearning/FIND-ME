// POST /api/ack — participant tapped "On my way".
import { send, readJson, cleanSeat, SEAT_RE, DEVICE_RE, postAppsScript, methodGuard } from './_lib.js';

export default async function handler(req, res) {
  if (!methodGuard(req, res, 'POST')) return;
  let body;
  try {
    body = await readJson(req);
  } catch {
    return send(res, 400, { ok: false, code: 'BAD_REQUEST' });
  }
  const seat = cleanSeat(body.seat);
  const deviceId = String(body.deviceId ?? '');
  if (!SEAT_RE.test(seat) || !DEVICE_RE.test(deviceId)) return send(res, 400, { ok: false, code: 'BAD_REQUEST' });

  try {
    const result = await postAppsScript({ action: 'ack', seat, deviceId, alertId: String(body.alertId ?? '').slice(0, 40) });
    return send(res, result.ok ? 200 : 409, result);
  } catch (err) {
    console.error('ack failed', err);
    return send(res, 502, { ok: false, code: 'UPSTREAM' });
  }
}
