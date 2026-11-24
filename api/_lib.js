// Shared helpers for the Find Me API (files starting with "_" are not deployed as endpoints).
import { timingSafeEqual } from 'node:crypto';

export const env = (key, fallback = '') => String(process.env[key] ?? fallback).trim();

export function send(res, status, data, headers = {}) {
  res.statusCode = status;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  if (!('Cache-Control' in headers)) res.setHeader('Cache-Control', 'no-store');
  for (const [k, v] of Object.entries(headers)) res.setHeader(k, v);
  res.end(JSON.stringify(data));
}

export async function readJson(req, limit = 64 * 1024) {
  if (req.body && typeof req.body === 'object' && !Buffer.isBuffer(req.body)) return req.body;
  if (typeof req.body === 'string') return JSON.parse(req.body || '{}');
  if (Buffer.isBuffer(req.body)) return JSON.parse(req.body.toString('utf8') || '{}');
  let raw = '';
  for await (const chunk of req) {
    raw += chunk;
    if (raw.length > limit) throw new Error('Body too large');
  }
  return JSON.parse(raw || '{}');
}

/** Seat labels like "A12" — upper-case, letters+digits only. */
export function cleanSeat(value) {
  return String(value ?? '').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 6);
}
export const SEAT_RE = /^[A-Z]{1,2}\d{1,3}$/;
export const DEVICE_RE = /^[A-Za-z0-9_-]{8,64}$/;

export function safeEqual(a, b) {
  const x = Buffer.from(String(a ?? ''));
  const y = Buffer.from(String(b ?? ''));
  return x.length > 0 && x.length === y.length && timingSafeEqual(x, y);
}

function appsScriptUrl() {
  const url = env('APPS_SCRIPT_URL');
  if (!url) throw new Error('APPS_SCRIPT_URL is not set in Vercel environment variables');
  return url;
}

async function parseAppsScript(r) {
  const text = await r.text();
  try {
    return JSON.parse(text);
  } catch {
    throw new Error(
      'Google Apps Script did not return JSON. Check the web app is deployed with access "Anyone". ' +
        `(HTTP ${r.status}: ${text.slice(0, 120).replace(/\s+/g, ' ')})`
    );
  }
}

/** POST an action to the Google Apps Script web app (secret added server-side). */
export async function postAppsScript(payload) {
  const r = await fetch(appsScriptUrl(), {
    method: 'POST',
    // text/plain avoids Apps Script content-type quirks; the body is JSON.
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ ...payload, secret: env('SHARED_SECRET') }),
    redirect: 'follow',
    signal: AbortSignal.timeout(20000),
  });
  return parseAppsScript(r);
}

/** GET an action from the Google Apps Script web app. */
export async function getAppsScript(params) {
  const url = new URL(appsScriptUrl());
  for (const [k, v] of Object.entries({ ...params, secret: env('SHARED_SECRET') })) url.searchParams.set(k, v);
  const r = await fetch(url, { redirect: 'follow', signal: AbortSignal.timeout(15000) });
  return parseAppsScript(r);
}

export function methodGuard(req, res, method) {
  if (req.method === method) return true;
  res.setHeader('Allow', method);
  send(res, 405, { ok: false, error: 'Method not allowed' });
  return false;
}
