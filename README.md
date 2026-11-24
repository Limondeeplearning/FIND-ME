# Find Me — seat alerts for the awarding ceremony

Participants scan a QR code, register their **name + seat**. When you select their seat on the Google Sheet seat map and click **📣 Find Me → Alert selected seats**, their phone **vibrates** (Android), **flashes full-screen** and shows:

> *Please proceed to the holding area on the right. Please find the staff to assist you.*

They tap **On my way** and the seat turns green on your sheet. Each participant's **name appears in the cell above their seat** on the seat map.

**Wrong-seat safety net:** before registering, participants see a small hall map with their seat highlighted and must confirm it. If a seat is already taken, they can tap **This is my seat — get staff help**, which turns the seat **red** on your map. Staff then unregister the wrong person. That phone gets a **"Please register again"** notification, and the person who asked gets the seat automatically.

```
Google Sheet (seat map + Apps Script)  ⇄  Vercel (web app + API)  →  participants' phones
```

| | Page open on screen | Phone locked / other app |
|---|---|---|
| Android | Vibrate + flash + chime | Notification + vibration |
| iPhone (added to Home Screen) | Flash + chime | Notification + vibration |
| iPhone (Safari tab only) | Flash + chime | ✖ — must keep page open |

---

## Setup (about 20 minutes)

### 1. Google Sheet — add the script
1. Open the sheet → **Extensions → Apps Script**.
2. Delete what's in `Code.gs`, paste all of [`apps-script/Code.gs`](apps-script/Code.gs), click **Save**.
3. Go back to the sheet and **reload** it. A **📣 Find Me** menu appears.
4. **📣 Find Me → 🔧 1. Set up / rebuild registry.** Approve the Google permission prompts.
   - Creates **Find Me Registry** (one row per seat, built from the seat map) and **Find Me Settings** (message, re-alert timing, registration on/off).
   - Re-run it any time the seat map changes. Existing registrations are kept.

### 2. Google Sheet — publish it as a web app
1. In Apps Script: **Deploy → New deployment → ⚙ Select type → Web app**.
2. **Execute as: Me** · **Who has access: Anyone** → **Deploy**.
3. Copy the **Web app URL** (ends in `/exec`). This is `APPS_SCRIPT_URL`.

> "Anyone" only lets the web address be called. Every call needs the secret password, and the sheet itself stays private.
> ⚠ When you change `Code.gs` later: **Deploy → Manage deployments → ✏ Edit → Version: New version → Deploy**, or the web app keeps the old code.

### 3. Push-notification keys
On any computer with Node.js, in this folder:
```bash
node scripts/generate-vapid-keys.mjs
```
Copy the `VAPID_PUBLIC_KEY` and `VAPID_PRIVATE_KEY` lines. Keep the private key secret.

### 4. Vercel — deploy
Push this folder to GitHub and import it in Vercel, or run `npx vercel` here. No build settings are needed.

In **Vercel → Project → Settings → Environment Variables**, add:

| Name | Value |
|---|---|
| `APPS_SCRIPT_URL` | Web app URL from step 2 |
| `VAPID_PUBLIC_KEY` | from step 3 |
| `VAPID_PRIVATE_KEY` | from step 3 |
| `VAPID_SUBJECT` | `mailto:your-email@example.com` |
| `EVENT_NAME` | `REBUILD Conference` |
| `EVENT_SUBTITLE` | `Day 2 · Awarding Proper` |
| `SHARED_SECRET` | from step 5 |

### 5. Connect the sheet ↔ Vercel
1. Sheet: **📣 Find Me → 🔗 2. Connect to Vercel**, then paste your Vercel address (e.g. `https://find-me.vercel.app`).
2. It shows a **SHARED_SECRET**. Paste it into Vercel as `SHARED_SECRET`, then **Redeploy** in Vercel.
3. Sheet: **📣 Find Me → 🧪 Test connection**. All four lines should be ✅.

### 6. QR code
Make a QR code of your Vercel address and print it on seat cards, the program, or the screen. In Chrome on a computer: open the site → **Share icon** in the address bar → **Create QR code** → **Download**. QR Code Monkey or Canva also work. Choose a **static** QR code, because some free "dynamic" ones stop working when their trial ends.

Tell participants to **scan with their Camera app**. If the link is opened inside Messenger, Facebook, Instagram, Viber and similar apps, Find Me shows an **Open in your browser** screen. On Android, a button switches to Chrome; on iPhone, it shows the steps. Registration is blocked there because those in-app browsers can't receive background alerts. On Xiaomi/Redmi/Poco, Oppo/Realme, vivo, Huawei, Infinix/Tecno and UC browsers, which phone QR scanners often open, it recommends Chrome with the same button and offers a **Continue anyway** link.

---

## During the event

| Do this | How |
|---|---|
| Call one or more people | Select their seat cells **or their name cells** on **ADMIN Seat Plan and Registration** (Ctrl/Cmd-click for several) → **📣 Find Me → 📣 Alert selected seats** |
| …or from the list | Tick **ALERT** in **Find Me Registry** |
| Cancel a call | Select seats → **↺ Cancel alert on selected seats** (or untick ALERT) |
| Change the message | Edit **Alert message** in **Find Me Settings**. It applies to the next alert. |
| See who hasn't responded | **📊 Status summary**, or look for orange seats |
| Stop new registrations | Untick **Registration open** in Settings |
| Fix a wrong registration | Select the seat → **✖ Unregister selected seats**. The phone gets a *"Please register again"* notification. |
| Red seat (someone says "this is my seat") | Check with an usher, then either **✖ Unregister** the wrong person (the claimant gets the seat automatically) or **⚠ Dismiss claim** if the claimant was wrong (their phone is asked to re-check their seat) |

**Seat map colours:** blue border = registered · **red** = needs staff (seat claimed by two people, or the same name at two seats; see the **⚠ Issue** column) · **orange** = alerted, waiting for reply (send an usher if it stays orange) · **green** = "On my way". Original seat colours come back when an alert is cancelled.

**Re-alerts:** seats still orange get re-sent about every minute, up to 3 times. Change this in Settings.

The toast after each alert lists seats that **aren't registered** (send an usher) and seats with **no lock-screen alerts** (they only see it while the page is open).

## After the event (Data Privacy)
**📣 Find Me → 🗑 Clear ALL registrations.** This deletes names and phone links and restores the map colours. Then set the sheet's sharing back to staff only.

---

## Updating the script later
After pasting a new `Code.gs`: **Deploy → Manage deployments → ✏ Edit → Version: New version → Deploy**, then run **📣 Find Me → 🔧 Set up / rebuild registry** once. This adds any new columns and keeps every registration.

## Rehearsal checklist
- [ ] Test on 1 Android + 1 iPhone (added to Home Screen): register → alert → locked-phone alert → On my way → green on the map
- [ ] Test the safety net: two phones register the same seat → **This is my seat** → seat turns red → unregister the wrong one → first phone gets *Please register again*, second gets *You're registered*
- [ ] Test on venue Wi-Fi **and** mobile data
- [ ] Usher announcement: *"Scan the QR, register your seat, tap Allow. iPhone: Share → Add to Home Screen first. Turn off Silent / Do Not Disturb."*
- [ ] Assign one staff member to run the sheet on a laptop

## Files
```
api/            Vercel serverless functions
  config.js     push key + event name for the phone
  register.js   phone → sheet: name + seat (+ push subscription); also "this is my seat" claims
  ack.js        phone → sheet: "On my way"
  alerts.js     phones poll the current calls + staff changes (cached 2 s at Vercel's edge)
  layout.js     seat positions for the "check your seat" hall map
  push.js       sheet → phones: sends the push notifications
public/         the participant web app (HTML/CSS/JS, service worker, icons)
apps-script/    Code.gs for the Google Sheet
scripts/        VAPID key generator
```

## Limits worth knowing
- iPhones can't vibrate from a web page. They flash and chime instead. Lock-screen notifications vibrate according to the phone's own settings.
- Do Not Disturb / Focus, Silent mode, battery saver or a force-closed browser can block notifications.
- Lock-screen alerts don't work in private/incognito tabs.
- The flash alternates once per second, which stays within the WCAG limit of 3 flashes per second.
