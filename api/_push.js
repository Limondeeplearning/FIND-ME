// Shared web-push sender (files starting with "_" are not deployed as endpoints).
import webpush from 'web-push';
import { env } from './_lib.js';

let ready = false;
export function initPush() {
  if (ready) return true;
  const publicKey = env('VAPID_PUBLIC_KEY');
  const privateKey = env('VAPID_PRIVATE_KEY');
  if (!publicKey || !privateKey) return false;
  webpush.setVapidDetails(env('VAPID_SUBJECT', 'mailto:admin@example.com'), publicKey, privateKey);
  ready = true;
  return true;
}

export const PUSH_TYPES = ['alert', 'unregistered', 'granted', 'dismissed', 'welcome'];

/** Sends one push. `payload` is { type, seat, alertId, device, title, message, serverTime }. */
export async function sendPush(subscription, payload) {
  if (!initPush()) throw Object.assign(new Error('VAPID keys are not set in Vercel'), { statusCode: 500 });
  const sub = typeof subscription === 'string' ? JSON.parse(subscription) : subscription;
  const body = JSON.stringify({
    type: PUSH_TYPES.includes(payload.type) ? payload.type : 'alert',
    seat: payload.seat,
    alertId: payload.alertId,
    device: payload.device || '',
    title: String(payload.title || '').slice(0, 80),
    message: String(payload.message || '').slice(0, 300),
    serverTime: Number(payload.serverTime) || Date.now(),
  });
  return webpush.sendNotification(sub, body, { TTL: 600, urgency: 'high', topic: 'findme' });
}
