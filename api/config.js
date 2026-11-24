// GET /api/config — public settings the phone needs (push key + event labels).
import { env, send, methodGuard } from './_lib.js';

export default function handler(req, res) {
  if (!methodGuard(req, res, 'GET')) return;
  send(
    res,
    200,
    {
      ok: true,
      vapidPublicKey: env('VAPID_PUBLIC_KEY'),
      eventName: env('EVENT_NAME', 'REBUILD Conference'),
      eventSubtitle: env('EVENT_SUBTITLE', 'Day 2 · Awarding Proper'),
    },
    { 'Cache-Control': 'public, max-age=0, s-maxage=60' }
  );
}
