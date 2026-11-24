/* Find Me — participant app
   Registers a seat (with a "check your seat" map), listens for alerts (web push + polling),
   runs the full-screen flash / vibration / chime until "On my way", and reacts to staff
   changes: unregistered → register again; "this is my seat" claim → seat granted. */
'use strict';

(() => {
  const DEFAULT_TITLE = "It's your turn";
  const DEFAULT_MSG = 'Please proceed to the holding area on the right. Please find the staff to assist you.';
  const POLL_MS = 4000;
  const CONFIRM_WAIT_MS = 25000; // how long to wait for the automatic "You're all set" push
  const KEYS = {
    reg: 'findme.reg', claim: 'findme.claim', device: 'findme.device', acked: 'findme.acked',
    awake: 'findme.awake', iosSkip: 'findme.iosSkip', lastName: 'findme.lastName',
  };
  const SEAT_RE = /^[A-Z]{1,2}\d{1,3}$/;

  const $ = (id) => document.getElementById(id);
  const store = {
    get(k, fallback = null) { try { const v = localStorage.getItem(k); return v === null ? fallback : JSON.parse(v); } catch { return fallback; } },
    set(k, v) { try { v === null ? localStorage.removeItem(k) : localStorage.setItem(k, JSON.stringify(v)); } catch { /* private mode */ } },
  };

  const ua = navigator.userAgent;
  const isIOS = /iPad|iPhone|iPod/.test(ua) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  const isStandalone = window.matchMedia('(display-mode: standalone)').matches || navigator.standalone === true;
  const pushSupported = 'serviceWorker' in navigator && 'PushManager' in window && 'Notification' in window;
  const canVibrate = typeof navigator.vibrate === 'function' && !isIOS;
  const wakeSupported = 'wakeLock' in navigator;
  const isAndroid = /Android/i.test(ua);

  /* In-app browsers (Messenger, Facebook, Instagram…) can't receive push alerts. */
  const IN_APP = [
    [/FB_IAB\/MESSENGER|MessengerForiOS|\bOrca-Android|FBAN\/Messenger/i, 'Messenger'],
    [/FBAN|FBAV|FB_IAB|FBIOS/i, 'Facebook'],
    [/Instagram/i, 'Instagram'],
    [/\bLine\//i, 'LINE'],
    [/Viber/i, 'Viber'],
    [/MicroMessenger/i, 'WeChat'],
    [/TikTok|musical_ly|Bytedance/i, 'TikTok'],
    [/Snapchat/i, 'Snapchat'],
    [/Telegram/i, 'Telegram'],
  ];
  /* Phone-maker browsers (often opened by the built-in QR scanner) — real browsers,
     but lock-screen alerts are unreliable there, so we recommend Chrome. */
  const VENDOR = [
    [/MiuiBrowser|XiaoMi\//i, 'Mi Browser'],
    [/HeyTapBrowser|OppoBrowser|RealmeBrowser/i, 'the phone browser'],
    [/VivoBrowser/i, 'vivo Browser'],
    [/HuaweiBrowser|HONORBrowser/i, 'the phone browser'],
    [/PHX\//i, 'Phoenix Browser'],
    [/UCBrowser|UCWEB/i, 'UC Browser'],
  ];
  function detectInApp() {
    for (const [re, name] of IN_APP) if (re.test(ua)) return { name, known: true };
    if (isAndroid) for (const [re, name] of VENDOR) if (re.test(ua)) return { name, known: false, vendor: true };
    if (isAndroid && /; wv\)/.test(ua)) return { name: 'another app', known: false };          // Android WebView
    if (isIOS && !isStandalone && !/Safari\//.test(ua)) return { name: 'another app', known: false }; // iOS WebView
    return null;
  }
  const inApp = detectInApp();

  const state = {
    config: { vapidPublicKey: '', eventName: '', eventSubtitle: '' },
    reg: store.get(KEYS.reg),       // { name, seat, endpoint, serverAt, regAt, verified, mobile }
    claim: store.get(KEYS.claim),   // { name, seat, endpoint, serverAt }
    pendingForm: null,              // { name, seat } between form and confirm
    layout: null,                   // { seats: Map(seat → [row, col]), stage, ... }
    devHash: '',
    missCount: 0,
    title: DEFAULT_TITLE,
    message: DEFAULT_MSG,
    current: null,                  // { alertId, test, acked }
    effectsTimer: null,
    pollTimer: null,
    wakeLock: null,
    audio: null,
    swReg: null,
    lastAckRetry: 0,
    syncTried: false,
    reachTimer: null,
    smsAvailable: store.get('findme.smsAvailable', false), // only offer texts when the sheet has SMS set up
  };

  /* ---------------- helpers ---------------- */
  function deviceId() {
    let id = store.get(KEYS.device);
    if (!id) {
      id = (crypto.randomUUID && crypto.randomUUID()) ||
        Array.from(crypto.getRandomValues(new Uint8Array(16)), (b) => b.toString(16).padStart(2, '0')).join('');
      store.set(KEYS.device, id);
    }
    return id;
  }
  async function computeDevHash() {
    try {
      const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(deviceId()));
      state.devHash = Array.from(new Uint8Array(buf).slice(0, 8), (b) => b.toString(16).padStart(2, '0')).join('');
    } catch { state.devHash = ''; }
  }
  const ackedSet = () => new Set(store.get(KEYS.acked, []));
  function markAcked(id) { store.set(KEYS.acked, [...ackedSet(), id].slice(-40)); }
  const normSeat = (v) => String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function show(name) {
    for (const s of ['loading', 'inapp', 'register', 'waiting', 'pending']) $('screen-' + s).hidden = s !== name;
    window.scrollTo(0, 0);
  }

  let toastTimer;
  function toast(msg, ms = 3500) {
    const t = $('toast');
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, ms);
  }

  function setConn(mode) {
    const pill = $('conn');
    pill.classList.toggle('live', mode === 'live' || mode === 'ready');
    pill.classList.toggle('down', mode === 'down');
    $('conn-text').textContent = { live: 'Live', ready: 'Ready', down: 'Reconnecting', idle: 'Connecting' }[mode] || mode;
  }

  async function api(path, body) {
    const opts = body
      ? { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) }
      : { method: 'GET' };
    const res = await fetch(path, opts);
    let data = {};
    try { data = await res.json(); } catch { /* empty */ }
    return { status: res.status, data };
  }

  function b64ToU8(b64) {
    const pad = '='.repeat((4 - (b64.length % 4)) % 4);
    const raw = atob((b64 + pad).replace(/-/g, '+').replace(/_/g, '/'));
    return Uint8Array.from(raw, (c) => c.charCodeAt(0));
  }

  function saveReg(reg) { state.reg = reg; store.set(KEYS.reg, reg); }
  function saveClaim(c) { state.claim = c; store.set(KEYS.claim, c); }

  /* ---------------- push ---------------- */
  async function getSubscription() {
    if (!pushSupported || Notification.permission !== 'granted' || !state.config.vapidPublicKey) return null;
    const reg = await navigator.serviceWorker.ready;
    const existing = await reg.pushManager.getSubscription();
    if (existing) return existing;
    return reg.pushManager.subscribe({ userVisibleOnly: true, applicationServerKey: b64ToU8(state.config.vapidPublicKey) });
  }

  /** Must be called synchronously inside a tap. Returns a promise for the subscription (or null). */
  function askPushInTap() {
    if (!pushSupported) return Promise.resolve(null);
    const perm = Notification.permission === 'default' ? Notification.requestPermission() : Promise.resolve(Notification.permission);
    return perm.then((p) => (p === 'granted' ? getSubscription() : null)).catch((err) => { console.warn('push subscribe failed', err); return null; });
  }

  function sendRegistration(name, seat, subscription, mode = 'register') {
    return api('/api/register', {
      mode,
      name,
      seat,
      deviceId: deviceId(),
      subscription: subscription ? subscription.toJSON() : null,
    });
  }

  /** Re-send registration silently if the push subscription changed. */
  async function syncSubscription() {
    if (!state.reg) return;
    try {
      const sub = await getSubscription();
      const endpoint = sub ? sub.endpoint : '';
      if (endpoint && endpoint !== state.reg.endpoint) {
        const { data } = await sendRegistration(state.reg.name, state.reg.seat, sub);
        if (data.ok) saveReg({ ...state.reg, endpoint, serverAt: data.serverTime || state.reg.serverAt, regAt: Date.now() - 2000, verified: false });
      }
    } catch (err) { console.warn('sync failed', err); }
    state.syncTried = true;
    updateChecks();
  }

  /* ---------------- audio + vibration ---------------- */
  function unlockAudio() {
    try {
      const AC = window.AudioContext || window.webkitAudioContext;
      if (!AC) return;
      if (!state.audio) state.audio = new AC();
      if (state.audio.state === 'suspended') state.audio.resume();
      const buf = state.audio.createBuffer(1, 1, 22050);
      const src = state.audio.createBufferSource();
      src.buffer = buf;
      src.connect(state.audio.destination);
      src.start(0);
    } catch { /* ignore */ }
    setTimeout(updateArm, 150);
  }

  function chime(notes = [880, 1175, 1568]) {
    const ctx = state.audio;
    if (!ctx || ctx.state !== 'running') return;
    const t0 = ctx.currentTime + 0.02;
    notes.forEach((f, i) => {
      const o = ctx.createOscillator();
      const g = ctx.createGain();
      o.type = 'triangle';
      o.frequency.value = f;
      const t = t0 + i * 0.2;
      g.gain.setValueAtTime(0.0001, t);
      g.gain.exponentialRampToValueAtTime(0.6, t + 0.02);
      g.gain.exponentialRampToValueAtTime(0.0001, t + 0.35);
      o.connect(g).connect(ctx.destination);
      o.start(t);
      o.stop(t + 0.4);
    });
  }

  function vibrate(pattern) {
    if (!canVibrate) return;
    try { navigator.vibrate(pattern); } catch { /* blocked until first tap */ }
  }

  function startEffects() {
    stopEffects();
    const pulse = () => { vibrate([700, 200, 700, 200, 1200]); chime(); };
    pulse();
    state.effectsTimer = setInterval(pulse, 3500);
  }
  function stopEffects() {
    clearInterval(state.effectsTimer);
    state.effectsTimer = null;
    vibrate(0);
  }

  /* ---------------- wake lock ---------------- */
  const awakePref = () => store.get(KEYS.awake, true);
  async function requestWake() {
    if (!wakeSupported || !awakePref() || document.hidden || state.wakeLock || !(state.reg || state.claim)) return;
    try {
      state.wakeLock = await navigator.wakeLock.request('screen');
      state.wakeLock.addEventListener('release', () => { state.wakeLock = null; });
    } catch { /* not allowed right now */ }
  }
  function releaseWake() {
    if (state.wakeLock) state.wakeLock.release().catch(() => {});
    state.wakeLock = null;
  }

  /* ---------------- alert overlay ---------------- */
  function triggerAlert(alertId, opts = {}) {
    if (!state.reg && !opts.test) return;
    if (state.current && state.current.alertId === alertId) return;
    state.current = { alertId, test: !!opts.test, acked: false };

    const el = $('alert');
    el.classList.remove('done');
    el.classList.toggle('test', !!opts.test);
    $('a-tag').textContent = opts.test ? 'Test alert' : 'Now calling';
    $('a-seat').textContent = state.reg ? state.reg.seat : '';
    $('alert-title').textContent = opts.test ? 'This is a test' : (opts.title || state.title || DEFAULT_TITLE);
    $('alert-msg').textContent = opts.test
      ? 'When it is your turn, your screen will flash like this. Tap the button to close.'
      : (opts.message || state.message || DEFAULT_MSG);
    const ack = $('ack-btn');
    ack.hidden = false;
    ack.disabled = false;
    ack.textContent = opts.test ? 'Close test' : 'On my way';
    $('alert-after').hidden = true;
    el.hidden = false;
    document.body.style.overflow = 'hidden';
    document.title = '🔔 ' + (opts.test ? 'Test alert' : "It's your turn") + ' · Find Me';
    // In the background the notification does the buzzing — don't fight it. Effects start when the page is seen.
    if (!document.hidden) startEffects();
    ack.focus({ preventScroll: true });
  }

  function closeAlert() {
    stopEffects();
    $('alert').hidden = true;
    document.body.style.overflow = '';
    document.title = 'Find Me · Seat Alerts';
    state.current = null;
  }

  async function sendAck(alertId) {
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        const { data } = await api('/api/ack', { seat: state.reg.seat, deviceId: deviceId(), alertId });
        if (data.ok) return true;
        if (data.code === 'NOT_YOUR_SEAT') return false;
      } catch { /* network — retry */ }
      await sleep(1200 * (attempt + 1));
    }
    return false;
  }

  async function onAck() {
    const cur = state.current;
    if (!cur) return;
    if (cur.test) return closeAlert();
    $('ack-btn').disabled = true;
    stopEffects();
    cur.acked = true;
    markAcked(cur.alertId);
    $('alert').classList.add('done');
    $('ack-btn').hidden = true;
    $('alert-after').hidden = false;
    $('close-btn').focus({ preventScroll: true });
    closeNotifications('findme-alert');
    const ok = await sendAck(cur.alertId);
    if (!ok) toast('Could not reach the desk — please proceed anyway.', 5000);
  }

  async function closeNotifications(tag) {
    try {
      const reg = state.swReg || (await navigator.serviceWorker.getRegistration());
      if (!reg) return;
      (await reg.getNotifications({ tag })).forEach((n) => n.close());
    } catch { /* ignore */ }
  }

  /* ---------------- staff notices ---------------- */
  const NOTICES = {
    unregistered: (seat) => ({
      tag: 'From the staff', title: 'Please register again', ok: false, btn: 'Register again',
      msg: `Staff have removed your registration${seat ? ' for seat ' + seat : ''}. Please register again with the seat number on your seat.`,
    }),
    dismissed: (seat) => ({
      tag: 'From the staff', title: 'Please check your seat', ok: false, btn: 'Register again',
      msg: `Staff checked seat ${seat || ''} and it belongs to someone else. Please register again with the seat number on your seat.`,
    }),
    granted: (seat) => ({
      tag: 'Seat confirmed', title: "You're registered", ok: true, btn: 'OK',
      msg: `Seat ${seat} is now yours. Stay seated — we'll alert you when it's your turn.`,
    }),
  };
  let noticeAction = null;

  function showNotice(kind, seat) {
    const n = NOTICES[kind](seat);
    const el = $('notice');
    el.classList.toggle('ok', n.ok);
    $('n-tag').textContent = n.tag;
    $('n-title').textContent = n.title;
    $('n-msg').textContent = n.msg;
    $('n-btn').textContent = n.btn;
    noticeAction = kind;
    el.hidden = false;
    document.body.style.overflow = 'hidden';
    vibrate([400, 150, 400]);
    chime(n.ok ? [880, 1320] : [660, 660]);
    $('n-btn').focus({ preventScroll: true });
  }

  function onNoticeButton() {
    $('notice').hidden = true;
    document.body.style.overflow = '';
    closeNotifications('findme-info');
    if (noticeAction === 'unregistered' || noticeAction === 'dismissed') {
      enterRegister();
      $('seat').value = '';
      $('seat').focus();
    }
    noticeAction = null;
  }

  /** Staff unregistered us / dismissed our claim → back to registration. */
  function onRemoved(kind) {
    const seat = (state.reg && state.reg.seat) || (state.claim && state.claim.seat) || '';
    if (state.reg) store.set(KEYS.lastName, state.reg.name);
    if (state.claim) store.set(KEYS.lastName, state.claim.name);
    saveReg(null);
    saveClaim(null);
    closeAlert();
    clearTimeout(state.pollTimer);
    releaseWake();
    state.missCount = 0;
    enterRegister();
    showNotice(kind, seat);
  }

  /** Our "this is my seat" claim was resolved in our favour. */
  function onGranted(seat, serverAt) {
    const c = state.claim || {};
    saveReg({ name: c.name || store.get(KEYS.lastName, ''), seat, endpoint: c.endpoint || '', serverAt: serverAt || Date.now(), regAt: Date.now() - CONFIRM_WAIT_MS, verified: false, mobile: '' });
    saveClaim(null);
    state.missCount = 0;
    enterWaiting();
    showNotice('granted', seat);
  }

  /* ---------------- polling ---------------- */
  async function poll() {
    clearTimeout(state.pollTimer);
    if (!state.reg && !state.claim) return;
    if (!document.hidden) {
      try {
        const { status, data } = await api('/api/alerts');
        if (status === 200 && data.ok) { setConn('live'); applyServer(data); if (state.reg && !state.reg.verified) checkReach(); }
        else setConn('down');
      } catch { setConn('down'); }
    }
    if (state.reg || state.claim) state.pollTimer = setTimeout(poll, POLL_MS + Math.random() * 1000);
  }

  function applyServer(data) {
    if (typeof data.smsAvailable === 'boolean' && data.smsAvailable !== state.smsAvailable) {
      state.smsAvailable = data.smsAvailable;
      store.set('findme.smsAvailable', data.smsAvailable);
      if (state.reg) setTimeout(checkReach, 0);
    }
    if (data.title) state.title = data.title;
    if (data.message) state.message = data.message;
    const devices = data.devices || null;
    const claims = data.claims || {};
    const gen = Number(data.generatedAt) || 0;
    const h = state.devHash;

    if (state.reg) {
      if (devices && h) {
        const seatNow = devices[h];
        if (!seatNow) {
          // Only trust "missing" if the sheet data is newer than our registration (twice in a row).
          if (gen > (state.reg.serverAt || 0)) {
            state.missCount += 1;
            if (state.missCount >= 2) return onRemoved('unregistered');
          }
          return;
        }
        state.missCount = 0;
        if (seatNow !== state.reg.seat) {
          saveReg({ ...state.reg, seat: seatNow });
          $('w-seat').textContent = seatNow;
          toast(`Staff updated your seat to ${seatNow}.`, 5000);
        }
      }
      const mine = (data.alerts || []).find((a) => a.seat === state.reg.seat);
      const acked = ackedSet();
      if (mine) {
        if (acked.has(mine.id)) {
          if (Date.now() - state.lastAckRetry > 15000) { state.lastAckRetry = Date.now(); sendAck(mine.id); }
        } else if (!state.current || state.current.alertId !== mine.id) {
          triggerAlert(mine.id);
        }
      } else if (state.current && !state.current.test && !state.current.acked) {
        // Cancelled / marked DONE in the sheet — but only trust data newer than this alert
        // (a slightly stale copy must not close an alert that a push just opened).
        const sentAt = parseInt(String(state.current.alertId).split('-')[0], 10);
        if (!sentAt || gen > sentAt) closeAlert();
      }
    } else if (state.claim && devices && h) {
      if (devices[h]) return onGranted(devices[h], gen);
      if (!claims[h] && gen > (state.claim.serverAt || 0)) {
        state.missCount += 1;
        if (state.missCount >= 2) return onRemoved('dismissed');
      } else state.missCount = 0;
    }
  }

  /* ---------------- reachability: push confirmed, or SMS fallback ---------------- */
  async function readPushReceipt() {
    try {
      const c = await caches.open('findme-state');
      const r = await c.match('/__pushok');
      return r ? (await r.json()).at || 0 : 0;
    } catch { return 0; }
  }

  const maskMobile = (m) => (m && m.length === 11 ? m.slice(0, 4) + ' *** ' + m.slice(7) : m || '');

  async function checkReach() {
    if (!state.reg) return;
    clearTimeout(state.reachTimer);
    const reg = state.reg;
    if (!reg.verified && reg.endpoint) {
      const at = await readPushReceipt();
      if (at && at >= (reg.regAt || 0) - 1000) {
        saveReg({ ...state.reg, verified: true });
        toast('Lock-screen alerts confirmed ✓');
      }
    }
    const r = state.reg;
    const permOk = pushSupported && Notification.permission === 'granted';
    const waited = Date.now() - (r.regAt || 0);
    const pending = !r.verified && permOk && r.endpoint && waited < CONFIRM_WAIT_MS;
    const needSms = !r.verified && !pending;
    if (pending) state.reachTimer = setTimeout(checkReach, CONFIRM_WAIT_MS - waited + 500);
    renderSms(needSms);
    updateChecks();
  }

  function renderSms(needSms) {
    const r = state.reg;
    const card = $('sms-card');
    const hasMobile = !!(r && r.mobile);
    if (!state.smsAvailable && !hasMobile) {
      // No SMS for this event: tell them to keep the page open / come back at awarding time.
      card.hidden = !needSms;
      card.classList.remove('saved');
      $('sms-form').hidden = true;
      $('sms-saved').hidden = true;
      $('sms-text').hidden = false;
      $('sms-kicker').textContent = 'Heads up';
      $('sms-title').textContent = "Notifications can't reach this phone";
      $('sms-text').textContent = "Please keep this page open, or come back here when it's awarding time. Your screen will flash and show your seat when you're called.";
      $('sms-link').hidden = true;
      return;
    }
    const show = hasMobile || needSms || card.dataset.open === '1';
    card.hidden = !show;
    card.classList.toggle('saved', hasMobile);
    $('sms-form').hidden = hasMobile && card.dataset.editing !== '1';
    $('sms-saved').hidden = !hasMobile || card.dataset.editing === '1';
    if (hasMobile) $('sms-masked').textContent = maskMobile(r.mobile);
    $('sms-text').hidden = hasMobile && card.dataset.editing !== '1';
    if (hasMobile && card.dataset.editing !== '1') {
      $('sms-kicker').textContent = 'Text alerts on';
      $('sms-title').textContent = "You're covered";
    } else if (needSms) {
      $('sms-kicker').textContent = 'Get your alert by text';
      $('sms-title').textContent = "Notifications can't reach this phone";
      $('sms-text').textContent = "Enter your mobile number and we'll text you when it's your turn. No settings needed.";
    } else {
      $('sms-kicker').textContent = 'Optional';
      $('sms-title').textContent = 'Also get a text message';
      $('sms-text').textContent = "We'll text you too when it's your turn.";
    }
    $('sms-link').hidden = show || !!(r && r.verified);
  }

  async function onSmsSubmit(ev) {
    ev.preventDefault();
    const err = $('sms-error');
    err.hidden = true;
    const raw = $('mobile').value;
    const digits = raw.replace(/[^0-9]/g, '');
    const mobile = digits.startsWith('639') ? '0' + digits.slice(2) : digits.length === 10 && digits.startsWith('9') ? '0' + digits : digits;
    if (raw.trim() && !/^09\d{9}$/.test(mobile)) {
      err.textContent = 'Please enter a PH mobile number like 0917 123 4567.';
      err.hidden = false;
      return;
    }
    setBusy('sms-save', true, 'Saving…');
    try {
      const { data } = await api('/api/register', { mode: 'mobile', seat: state.reg.seat, deviceId: deviceId(), mobile: raw.trim() ? mobile : '' });
      if (!data.ok) {
        err.textContent = data.code === 'MOBILE_INVALID' ? 'Please enter a PH mobile number like 0917 123 4567.'
          : data.code === 'NOT_YOUR_SEAT' ? 'This phone is no longer registered to that seat. Please register again.'
          : 'Could not save right now. Please try again.';
        err.hidden = false;
        return;
      }
      saveReg({ ...state.reg, mobile: data.mobile || '' });
      $('sms-card').dataset.editing = '';
      $('sms-note').textContent = !data.mobile ? '' : data.confirmSent ? 'We just sent you a confirmation text.'
        : data.smsReady ? '' : 'Please also keep this page open, just in case.';
      toast(data.mobile ? 'Mobile number saved.' : 'Mobile number removed.');
      checkReach();
    } catch {
      err.textContent = "Can't connect. Check your Wi-Fi or mobile data and try again.";
      err.hidden = false;
    } finally {
      setBusy('sms-save', false);
    }
  }

  /* ---------------- status checklist ---------------- */
  function setChk(id, level, text) {
    const li = $(id);
    li.classList.remove('ok', 'warn', 'bad');
    if (level) li.classList.add(level);
    li.querySelector('small').textContent = text;
  }

  function updateChecks() {
    const fallback = state.smsAvailable ? 'add your mobile number above' : "keep this page open or come back here when it's awarding time";
    const btn = $('notif-btn');
    btn.hidden = true;
    if (!pushSupported) {
      if (state.reg && state.reg.mobile) setChk('chk-notif', 'warn', "Not available here — you'll get a text instead.");
      else if (isIOS && !isStandalone) setChk('chk-notif', 'warn', 'Only while this page is open. Add Find Me to your Home Screen, or ' + fallback + '.');
      else setChk('chk-notif', 'warn', 'Not available on this browser — keep this page open on screen.');
    } else if (Notification.permission === 'granted') {
      const r = state.reg || {};
      if (r.endpoint && r.verified) setChk('chk-notif', 'ok', 'Confirmed — alerts reach you even when your phone is locked.');
      else if (r.endpoint && Date.now() - (r.regAt || 0) < CONFIRM_WAIT_MS) setChk('chk-notif', 'warn', 'Checking… you should get a "You\'re all set" notification in a few seconds.');
      else if (r.endpoint) setChk('chk-notif', 'warn', r.mobile ? 'Not confirmed on this phone — you\'ll get a text instead.' : 'Not confirmed — this phone may block notifications. Please ' + fallback + '.');
      else if (state.syncTried) setChk('chk-notif', 'warn', "This browser couldn't turn on lock-screen alerts (private/incognito tab?). Keep this page open.");
      else setChk('chk-notif', 'warn', 'Allowed — finishing setup…');
    } else if (Notification.permission === 'denied') {
      if (state.reg && state.reg.mobile) setChk('chk-notif', 'warn', "Off on this phone — you'll get a text instead.");
      else setChk('chk-notif', 'bad', 'Blocked on this phone. Please ' + fallback + '.');
    } else {
      setChk('chk-notif', 'warn', state.reg && state.reg.mobile ? "Off — you'll get a text instead." : 'Off — tap Turn on, or ' + fallback + '.');
      btn.hidden = false;
    }

    if (canVibrate) setChk('chk-vibe', 'ok', 'On — your phone will buzz.');
    else if (isIOS) setChk('chk-vibe', 'warn', 'iPhone: the screen flashes and chimes instead. Lock-screen alerts still vibrate.');
    else setChk('chk-vibe', 'warn', 'Not supported here — the screen will flash instead.');

    const toggle = $('awake-toggle');
    if (wakeSupported) {
      toggle.checked = awakePref();
      toggle.closest('.switch').hidden = false;
      setChk('chk-awake', toggle.checked ? 'ok' : 'warn', toggle.checked ? 'On — your screen stays on while this page is open.' : 'Off — your screen may sleep.');
    } else {
      toggle.closest('.switch').hidden = true;
      setChk('chk-awake', 'warn', 'Not supported — set Auto-Lock to Never or keep tapping the screen.');
    }
    updateArm();
  }

  function updateArm() {
    if (!state.reg) return;
    const audioReady = state.audio && state.audio.state === 'running';
    const activated = navigator.userActivation ? navigator.userActivation.hasBeenActive : audioReady;
    $('arm-btn').hidden = !!(audioReady && activated);
  }

  /* ---------------- hall map (confirm step) ---------------- */
  async function loadLayout() {
    try {
      const { status, data } = await api('/api/layout');
      if (status !== 200 || !data.ok || !Array.isArray(data.seats) || !data.seats.length) return;
      const seats = new Map(data.seats.map(([s, r, c]) => [s, [r, c]]));
      const cols = [...new Set(data.seats.map((x) => x[2]))].sort((a, b) => a - b);
      const rows = [...new Set(data.seats.map((x) => x[1]))].sort((a, b) => a - b);
      // Split columns into sections wherever there's an aisle (gap) between them.
      const blocks = [];
      cols.forEach((c, i) => { if (i === 0 || c - cols[i - 1] > 1) blocks.push([]); blocks[blocks.length - 1].push(c); });
      const names = blocks.length === 3 ? ['Left', 'Center', 'Right'] : blocks.length === 2 ? ['Left', 'Right'] : blocks.map((_, i) => 'Section ' + (i + 1));
      const blockOf = new Map();
      blocks.forEach((b, i) => b.forEach((c) => blockOf.set(c, i)));
      state.layout = { seats, cols, rows, blocks, names, blockOf, stage: data.stage };
    } catch { /* map is optional */ }
  }

  function describeSeat(seat) {
    const L = state.layout;
    const rowLetter = seat.replace(/\d+$/, '');
    const num = seat.replace(/^[A-Z]+/, '');
    let where = `Row ${rowLetter} · Seat ${num}`;
    if (L && L.seats.has(seat) && L.blocks.length > 1) {
      const block = L.blockOf.get(L.seats.get(seat)[1]);
      where = `Row ${rowLetter} · ${L.names[block]} section · Seat ${num}`;
    }
    return where;
  }

  function renderHall(seat) {
    const L = state.layout;
    const box = $('hall');
    box.textContent = '';
    if (!L || !L.seats.has(seat)) { box.hidden = true; return; }
    box.hidden = false;
    const [selR] = L.seats.get(seat);
    const byPos = new Map();
    L.seats.forEach(([r, c], s) => byPos.set(r + ':' + c, s));
    const stageTop = L.stage && L.stage[0] < L.rows[0];
    const stageBottom = L.stage && L.stage[0] > L.rows[L.rows.length - 1];
    const stage = () => { const d = document.createElement('div'); d.className = 'hall-stage' + (stageBottom ? ' bottom' : ''); d.textContent = 'STAGE'; return d; };
    if (stageTop) box.appendChild(stage());
    const grid = document.createElement('div');
    grid.className = 'hall-grid';
    L.rows.forEach((r) => {
      const rowEl = document.createElement('div');
      rowEl.className = 'hall-row' + (r === selR ? ' sel-row' : '');
      const first = L.cols.map((c) => byPos.get(r + ':' + c)).find(Boolean) || '';
      const lbl = document.createElement('span');
      lbl.className = 'rl';
      lbl.textContent = first.replace(/\d+$/, '');
      rowEl.appendChild(lbl);
      L.cols.forEach((c, i) => {
        if (i > 0 && L.blockOf.get(c) !== L.blockOf.get(L.cols[i - 1])) {
          const g = document.createElement('span'); g.className = 'g'; rowEl.appendChild(g);
        }
        const s = byPos.get(r + ':' + c);
        const cell = document.createElement('span');
        cell.className = s ? 's' + (s === seat ? ' sel' : '') : 'g';
        rowEl.appendChild(cell);
      });
      grid.appendChild(rowEl);
    });
    box.appendChild(grid);
    if (stageBottom) box.appendChild(stage());
  }

  /* ---------------- opened inside another app ---------------- */
  function enterInApp() {
    show('inapp');
    setConn('ready');
    const url = location.origin + '/';
    if (inApp.vendor) {
      const lede = $('ia-lede');
      lede.textContent = '';
      const b1 = document.createElement('b'); b1.textContent = inApp.name;
      const b2 = document.createElement('b'); b2.textContent = 'Chrome';
      lede.append('This page opened in ', b1, '. Lock-screen alerts may not reach you there, so please open Find Me in ', b2, '.');
      $('ia-continue').textContent = 'Continue in ' + inApp.name + ' anyway (you may miss alerts)';
    } else {
      $('ia-app').textContent = inApp.name;
    }
    $('ia-url').value = url;
    const steps = $('ia-steps');
    steps.textContent = '';
    const add = (html) => {
      const li = document.createElement('li');
      const n = document.createElement('span');
      n.className = 'n';
      n.textContent = String(steps.children.length + 1);
      const t = document.createElement('span');
      t.innerHTML = html; // static strings only
      li.append(n, t);
      steps.appendChild(li);
    };
    if (isAndroid) {
      if (!inApp.vendor) $('ia-browser').textContent = 'Chrome';
      const open = $('ia-open');
      open.href = 'intent://' + location.host + '/#Intent;scheme=https;package=com.android.chrome;S.browser_fallback_url=' +
        encodeURIComponent(url) + ';end';
      open.hidden = false;
      $('ia-kicker').textContent = "If the button doesn't work";
      if (inApp.vendor) {
        add('Open the <b>Chrome</b> app.');
        add('Tap <b>Copy</b> below, then paste the link into Chrome\'s address bar.');
        add('Tip: next time, scan with <b>Google Lens</b> (Google app or Chrome\'s camera icon) to open in Chrome directly.');
      } else {
        add('Tap the <span class="kbd">⋮</span> menu at the top right.');
        add('Choose <b>Open in Chrome</b> or <b>Open in browser</b>.');
      }
    } else {
      $('ia-browser').textContent = 'Safari';
      $('ia-open').hidden = true;
      $('ia-kicker').textContent = 'How to open it';
      add('Tap <span class="kbd">•••</span> at the top right, or the <b>Share</b> / compass icon at the bottom.');
      add('Choose <b>Open in Safari</b> (or <b>Open in browser</b>).');
      add('In Safari, tap <b>Share → Add to Home Screen</b>, then open Find Me from your Home Screen.');
    }
    // Unknown web views might be a real browser we didn't recognise — allow a way through.
    $('ia-continue').hidden = inApp.known;
  }

  async function copyLink() {
    const input = $('ia-url');
    try {
      await navigator.clipboard.writeText(input.value);
      toast('Link copied. Paste it into your browser.');
    } catch {
      input.focus();
      input.select();
      try { document.execCommand('copy'); toast('Link copied. Paste it into your browser.'); }
      catch { toast('Press and hold the link to copy it.'); }
    }
  }

  /* ---------------- registration flow ---------------- */
  function showStep(step) {
    $('reg-form').hidden = step !== 'form';
    $('confirm').hidden = step !== 'confirm';
    $('taken').hidden = step !== 'taken';
    if (step !== 'form') window.scrollTo(0, 0);
  }

  function enterRegister() {
    clearTimeout(state.pollTimer);
    releaseWake();
    show('register');
    showStep('form');
    setConn('ready');
    $('ios-guide').hidden = !(isIOS && !isStandalone && !store.get(KEYS.iosSkip, false));
    const last = store.get(KEYS.lastName, '');
    if (last && !$('name').value) $('name').value = last;
  }

  function enterWaiting() {
    show('waiting');
    $('w-seat').textContent = state.reg.seat;
    $('w-name').textContent = state.reg.name;
    $('sms-card').dataset.open = '';
    $('sms-card').dataset.editing = '';
    $('sms-note').textContent = '';
    updateChecks();
    requestWake();
    poll();
    syncSubscription();
    checkReach();
  }

  function enterPending() {
    show('pending');
    $('p-seat').textContent = state.claim.seat;
    requestWake();
    poll();
  }

  /** Step 1: name + seat typed → validate, then show the hall map to confirm. */
  function onRegisterForm(ev) {
    ev.preventDefault();
    unlockAudio();
    const errEl = $('reg-error');
    errEl.hidden = true;
    const name = $('name').value.replace(/\s+/g, ' ').trim();
    const seat = normSeat($('seat').value);
    const fail = (msg, field) => { errEl.textContent = msg; errEl.hidden = false; if (field) $(field).focus(); };

    if (name.length < 2) return fail('Please enter your full name.', 'name');
    if (!SEAT_RE.test(seat)) return fail('Please enter your seat number, like A12.', 'seat');
    if (state.layout && !state.layout.seats.has(seat)) {
      return fail(`We couldn't find seat ${seat}. Please check the number on your seat or ask a staff member.`, 'seat');
    }
    state.pendingForm = { name, seat };
    $('c-seat').textContent = seat;
    $('c-where').textContent = describeSeat(seat);
    $('c-error').hidden = true;
    renderHall(seat);
    showStep('confirm');
    $('confirm-yes').focus({ preventScroll: true });
  }

  function setBusy(btnId, busy, label) {
    const b = $(btnId);
    b.disabled = busy;
    const l = b.querySelector('.btn-label') || b;
    if (!b.dataset.label) b.dataset.label = l.textContent;
    l.textContent = busy ? label : b.dataset.label;
  }

  const ERRORS = {
    SEAT_NOT_FOUND: (s) => `We couldn't find seat ${s}. Please check the number on your seat or ask a staff member.`,
    CLOSED: () => 'Registration is closed. Please ask a staff member.',
    NAME_REQUIRED: () => 'Please enter your full name.',
    SEAT_INVALID: () => 'Please enter your seat number, like A12.',
    ERROR: () => 'The registration desk is busy. Please try again in a few seconds.',
    CLAIM_EXISTS: () => 'Someone else has already asked staff about this seat. Please stay seated — a staff member will come to you.',
  };

  async function postWithRetry(name, seat, sub, mode) {
    let { status, data } = await sendRegistration(name, seat, sub, mode);
    if (!data.ok && (data.code === 'ERROR' || status >= 500)) {
      await sleep(2500 + Math.random() * 2000); // sheet busy — retry once
      ({ status, data } = await sendRegistration(name, seat, sub, mode));
    }
    return { status, data };
  }

  function registeredOK(name, seat, sub, data, t0) {
    store.set(KEYS.lastName, name);
    saveClaim(null);
    saveReg({ name, seat, endpoint: sub ? sub.endpoint : '', serverAt: data.serverTime || 0, regAt: t0 || Date.now(), verified: false, mobile: '' });
    state.missCount = 0;
    enterWaiting();
    toast(sub ? 'You are registered. Checking alerts…' : 'You are registered.');
    if (data.status === 'Alerted' && data.alertId && !ackedSet().has(data.alertId)) triggerAlert(data.alertId);
  }

  /** Step 2: "Yes, that's my seat". */
  async function onConfirmYes() {
    const f = state.pendingForm;
    if (!f) return showStep('form');
    const subPromise = askPushInTap(); // must start inside the tap
    unlockAudio();
    setBusy('confirm-yes', true, 'Registering…');
    $('c-error').hidden = true;
    try {
      const sub = await subPromise;
      const t0 = Date.now();
      const { status, data } = await postWithRetry(f.name, f.seat, sub, 'register');
      if (data.ok) return registeredOK(f.name, f.seat, sub, data, t0);
      if (data.code === 'SEAT_TAKEN') {
        $('t-seat').textContent = f.seat;
        $('t-error').hidden = true;
        showStep('taken');
        return;
      }
      const msg = (ERRORS[data.code] || (() => (status >= 500 ? "Can't reach the registration desk right now. Please try again." : 'Something went wrong. Please try again.')))(f.seat);
      $('c-error').textContent = msg;
      $('c-error').hidden = false;
    } catch {
      $('c-error').textContent = "Can't connect. Check your Wi-Fi or mobile data and try again.";
      $('c-error').hidden = false;
    } finally {
      setBusy('confirm-yes', false);
    }
  }

  /** Seat taken → "This is my seat — get staff help". */
  async function onClaim() {
    const f = state.pendingForm;
    if (!f) return showStep('form');
    const subPromise = askPushInTap();
    setBusy('taken-claim', true, 'Notifying staff…');
    $('t-error').hidden = true;
    try {
      const sub = await subPromise;
      const t0 = Date.now();
      const { data } = await postWithRetry(f.name, f.seat, sub, 'claim');
      if (data.ok && data.registered) return registeredOK(f.name, f.seat, sub, data, t0);
      if (data.ok && data.claimed) {
        store.set(KEYS.lastName, f.name);
        saveClaim({ name: f.name, seat: f.seat, endpoint: sub ? sub.endpoint : '', serverAt: data.serverTime || 0 });
        state.missCount = 0;
        enterPending();
        return;
      }
      $('t-error').textContent = (ERRORS[data.code] || (() => 'Something went wrong. Please try again or ask a staff member.'))(f.seat);
      $('t-error').hidden = false;
    } catch {
      $('t-error').textContent = "Can't connect. Check your Wi-Fi or mobile data and try again.";
      $('t-error').hidden = false;
    } finally {
      setBusy('taken-claim', false);
    }
  }

  async function onUnclaim() {
    if (!window.confirm('Go back and enter a different seat number?')) return;
    const c = state.claim;
    saveClaim(null);
    if (c) api('/api/register', { mode: 'unclaim', seat: c.seat, deviceId: deviceId() }).catch(() => {});
    enterRegister();
    $('seat').value = '';
    $('seat').focus();
  }

  async function onEnableNotifications() {
    unlockAudio();
    try {
      const p = await Notification.requestPermission();
      if (p === 'granted') {
        const sub = await getSubscription();
        if (sub) {
          const { data } = await sendRegistration(state.reg.name, state.reg.seat, sub);
          if (data.ok) { saveReg({ ...state.reg, endpoint: sub.endpoint, serverAt: data.serverTime || state.reg.serverAt, regAt: Date.now() - 2000, verified: false }); toast('Lock-screen alerts are on — checking…'); checkReach(); }
        }
      } else toast('Notifications were not allowed.');
    } catch { toast('Could not turn on notifications.'); }
    updateChecks();
  }

  async function onTest() {
    unlockAudio();
    triggerAlert('test-' + Date.now(), { test: true });
    try {
      if (pushSupported && Notification.permission === 'granted') {
        const reg = await navigator.serviceWorker.ready;
        reg.showNotification('Find Me · Test alert', {
          body: 'This is how your lock-screen alert will look.',
          tag: 'findme-test',
          icon: '/icons/icon-192.png',
          badge: '/icons/badge-96.png',
          vibrate: [500, 200, 500],
        });
      }
    } catch { /* ignore */ }
  }

  function onChangeSeat() {
    if (!window.confirm('Change your registered seat?')) return;
    saveReg(null);
    closeAlert();
    $('seat').value = '';
    enterRegister();
    $('seat').focus();
  }

  /* ---------------- init ---------------- */
  async function init() {
    $('reg-form').addEventListener('submit', onRegisterForm);
    $('seat').addEventListener('input', (e) => {
      const v = normSeat(e.target.value).slice(0, 5);
      if (e.target.value !== v) e.target.value = v;
    });
    $('ios-skip').addEventListener('click', () => { store.set(KEYS.iosSkip, true); $('ios-guide').hidden = true; $('name').focus(); });
    $('confirm-yes').addEventListener('click', onConfirmYes);
    $('confirm-edit').addEventListener('click', () => { showStep('form'); $('seat').focus(); $('seat').select(); });
    $('taken-edit').addEventListener('click', () => { showStep('form'); $('seat').value = ''; $('seat').focus(); });
    $('taken-claim').addEventListener('click', onClaim);
    $('unclaim-btn').addEventListener('click', onUnclaim);
    $('ack-btn').addEventListener('click', onAck);
    $('close-btn').addEventListener('click', closeAlert);
    $('n-btn').addEventListener('click', onNoticeButton);
    $('ia-copy').addEventListener('click', copyLink);
    $('sms-form').addEventListener('submit', onSmsSubmit);
    $('sms-change').addEventListener('click', () => {
      const card = $('sms-card');
      card.dataset.editing = '1';
      $('mobile').value = (state.reg && state.reg.mobile) || '';
      renderSms(false);
      $('mobile').focus();
    });
    $('sms-link').addEventListener('click', () => { $('sms-card').dataset.open = '1'; renderSms(false); $('mobile').focus(); });
    $('ia-continue').addEventListener('click', () => { store.set('findme.iaSkip', true); enterRegister(); });
    $('test-btn').addEventListener('click', onTest);
    $('change-btn').addEventListener('click', onChangeSeat);
    $('notif-btn').addEventListener('click', onEnableNotifications);
    $('arm-btn').addEventListener('click', () => { unlockAudio(); vibrate(60); setTimeout(updateArm, 200); toast('Sound and vibration are armed.'); });
    $('awake-toggle').addEventListener('change', (e) => {
      store.set(KEYS.awake, e.target.checked);
      if (e.target.checked) requestWake(); else releaseWake();
      updateChecks();
    });
    document.addEventListener('pointerdown', unlockAudio, { once: true, capture: true });

    document.addEventListener('visibilitychange', () => {
      document.body.classList.toggle('paused', document.hidden);
      if (document.hidden) { if (state.effectsTimer) stopEffects(); return; }
      if (state.current && !state.current.acked && !state.effectsTimer) startEffects(); // alert waiting for them
      if (state.reg || state.claim) { requestWake(); poll(); if (state.reg) { updateChecks(); checkReach(); } }
    });

    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('/sw.js').then((r) => { state.swReg = r; }).catch((e) => console.warn('SW failed', e));
      navigator.serviceWorker.addEventListener('message', (e) => {
        const d = e.data || {};
        if (d.type === 'findme-alert' && d.alertId && state.reg && !ackedSet().has(d.alertId)) {
          if (d.title) state.title = d.title;
          if (d.message) state.message = d.message;
          triggerAlert(d.alertId, { title: d.title, message: d.message });
        }
        if (d.type === 'findme-notice') {
          if (d.kind === 'unregistered' && state.reg && state.reg.seat === d.seat) onRemoved('unregistered');
          else if (d.kind === 'dismissed' && state.claim && state.claim.seat === d.seat) onRemoved('dismissed');
          else if (d.kind === 'granted' && state.claim && state.claim.seat === d.seat) onGranted(d.seat, d.serverTime);
          else if (d.kind === 'welcome' && state.reg) setTimeout(checkReach, 300);
        }
      });
    }

    loadLayout();
    try {
      const { data } = await api('/api/config');
      if (data.ok) state.config = data;
    } catch { /* offline — use defaults */ }
    if (state.config.eventName) $('event-name').textContent = state.config.eventName;
    if (state.config.eventSubtitle) $('event-sub').textContent = state.config.eventSubtitle;
    await computeDevHash();

    const params = new URLSearchParams(location.search);
    const fromNotif = params.get('alert');
    if (fromNotif) history.replaceState(null, '', location.pathname);

    if (state.reg) {
      enterWaiting();
      if (fromNotif && !ackedSet().has(fromNotif)) triggerAlert(fromNotif);
    } else if (state.claim) {
      enterPending();
    } else if (inApp && !store.get('findme.iaSkip', false)) {
      enterInApp();
    } else {
      enterRegister();
    }
  }

  init();
})();
