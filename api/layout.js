// GET /api/layout — seat positions for the hall map shown on the "check your seat" step.
import { send, getAppsScript, methodGuard } from './_lib.js';

export default async function handler(req, res) {
  if (!methodGuard(req, res, 'GET')) return;
  try {
    const data = await getAppsScript({ action: 'layout' });
    if (!data.ok) return send(res, 502, { ok: false });
    return send(res, 200, { ok: true, seats: data.seats || [], stage: data.stage || null }, {
      'Cache-Control': 'public, max-age=60, s-maxage=300, stale-while-revalidate=600',
    });
  } catch (err) {
    console.error('layout failed', err);
    return send(res, 502, { ok: false });
  }
}
