// GET /api/alerts — the list of seats currently being called.
// Every phone polls this while its screen is on. Vercel's CDN caches it for 2 s,
// so the Google Sheet is asked at most about once every 2 s no matter how many phones poll.
import { send, getAppsScript, methodGuard } from './_lib.js';

export default async function handler(req, res) {
  if (!methodGuard(req, res, 'GET')) return;
  try {
    const data = await getAppsScript({ action: 'alerts' });
    if (!data.ok) return send(res, 502, { ok: false, error: data.error || 'Sheet error' });
    return send(
      res,
      200,
      {
        ok: true,
        title: data.title,
        message: data.message,
        registrationOpen: data.registrationOpen !== false,
        alerts: Array.isArray(data.alerts) ? data.alerts : [],
        devices: data.devices || {},
        claims: data.claims || {},
        generatedAt: data.generatedAt || 0,
      },
      { 'Cache-Control': 'public, max-age=0, s-maxage=2, stale-while-revalidate=4' }
    );
  } catch (err) {
    console.error('alerts failed', err);
    return send(res, 502, { ok: false, error: 'Could not reach the sheet' });
  }
}
