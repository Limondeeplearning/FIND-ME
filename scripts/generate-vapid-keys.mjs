// Generates the push-notification key pair (VAPID) — no packages needed.
// Run:  node scripts/generate-vapid-keys.mjs
// Then paste both values into Vercel → Project → Settings → Environment Variables.
import { generateKeyPairSync } from 'node:crypto';

const { publicKey, privateKey } = generateKeyPairSync('ec', { namedCurve: 'prime256v1' });
const pub = publicKey.export({ format: 'jwk' });
const priv = privateKey.export({ format: 'jwk' });

const b64u = (s) => Buffer.from(s, 'base64url');
const rawPublic = Buffer.concat([Buffer.from([0x04]), b64u(pub.x), b64u(pub.y)]).toString('base64url');

console.log('\nVAPID_PUBLIC_KEY=' + rawPublic);
console.log('VAPID_PRIVATE_KEY=' + priv.d + '\n');
console.log('Keep the private key secret. Do not put it in the sheet or in code.\n');
