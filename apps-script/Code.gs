/**
 * FIND ME — Google Sheet controller (Google Apps Script)
 * ------------------------------------------------------------------
 * Paste this whole file into Extensions → Apps Script → Code.gs.
 *
 * What it does
 *  • Adds a "📣 Find Me" menu to the sheet.
 *  • Builds a "Find Me Registry" tab from the seat labels on your seat map.
 *  • Phones register (name + seat) → the name fills in automatically.
 *  • Select seats on the seat map → 📣 Find Me → Alert selected seats.
 *    (or tick the ALERT checkbox in the registry)
 *  • Seat map changes live: blue border = registered, orange = alerted,
 *    green = "On my way". Original seat colours come back on reset.
 *  • Re-alerts seats that have not responded (every ~1 min, max 3 times).
 *  • Safety net: if someone says "This is my seat" for a taken seat, the seat
 *    turns RED with a note. Staff unregister the wrong person → that phone is
 *    told to register again, and the person who asked gets the seat automatically.
 *  • Same name registered at two seats → both flagged red.
 *  • DONE ☑ (award received) greys the seat and blocks alerts; untick to call again.
 *  • UNREGISTER ☑ in the registry = same as the Unregister menu, per seat.
 *  • Phones that can't get push notifications can give a mobile number → SMS
 *    via Semaphore (📣 Find Me → 📱 Set up SMS). "Alert via" shows how each
 *    seat will be reached: Push ✓ (confirmed), SMS, or Page only.
 *  • Each participant's name is written in the cell directly above their seat
 *    on the seat map (e.g. row 9 for row-A seats in row 10). Selecting the
 *    name cell works the same as selecting the seat.
 */

/* ======================= CONSTANTS ======================= */
var REG_SHEET = 'Find Me Registry';
var SET_SHEET = 'Find Me Settings';
var DEFAULT_MAP = 'ADMIN Seat Plan and Registration';
var SEAT_RE = /^[A-Z]{1,2}\d{1,3}$/;

// Registry columns (1-based)
var C = {
  SEAT: 1, NAME: 2, REGISTERED: 3, ALERT: 4, DONE: 5, UNREG: 6, STATUS: 7, VIA: 8,
  ALERTED_AT: 9, ACK_AT: 10, COUNT: 11, ISSUE: 12,
  // hidden from here on
  MOBILE: 13, MAP_CELL: 14, DEVICE: 15, SUB: 16, ALERT_ID: 17, LAST_SENT: 18, ORIG_BG: 19,
  CLAIM: 20, NAME_CELL: 21, PUSH_OK: 22
};
var NCOLS = 22;
// Header text is used to carry data over when columns move between versions — keep names stable.
var HEADERS = ['Seat', 'Name', 'Registered', 'ALERT', 'DONE', 'UNREGISTER', 'Status', 'Alert via',
  'First alerted', 'On my way at', 'Alerts sent', '⚠ Issue',
  'Mobile (hidden)', 'Map cell', 'Device', 'Push subscription', 'Alert ID', 'Last sent', 'Original colour',
  'Claim (hidden)', 'Name cell (hidden)', 'Push confirmed (hidden)'];

var STATUS = { NONE: '', REGISTERED: 'Registered', ALERTED: 'Alerted', ON_THE_WAY: 'On the way', DONE: 'Done ✓' };
var COLORS = {
  registeredBorder: '#2563EB',
  alertedFill: '#FDBA74', alertedBorder: '#C2410C',
  onWayFill: '#86EFAC', onWayBorder: '#15803D',
  issueFill: '#FCA5A5', issueBorder: '#B91C1C',
  doneFill: '#D1D5DB', doneBorder: '#4B5563',
  plainBorder: '#BDBDBD'
};

var NOTICE = {
  unregistered: { title: 'Please register again', message: 'Staff have removed your seat registration. Tap here to register with the seat number on your seat.' },
  granted: { title: "You're registered", message: 'Seat {seat} is now yours. Stay seated — we will alert you when it is your turn.' },
  dismissed: { title: 'Please check your seat', message: 'Staff checked seat {seat}. Please register again with the seat number on your seat.' }
};

var DEFAULT_SETTINGS = [
  ['Alert title', "It's your turn", 'Big heading on the phone screen and notification title'],
  ['Alert message', 'Please proceed to the holding area on the right. Please find the staff to assist you.', 'Shown on the phone. You can change this live.'],
  ['Seat map tab', DEFAULT_MAP, 'Exact name of the seat map tab'],
  ['Re-alert after (seconds)', 60, 'Resend to seats that have not tapped "On my way" (checked about once a minute; 0 = off)'],
  ['Max alerts per seat', 3, 'Stop re-alerting after this many sends'],
  ['Registration open', true, 'Untick to stop new registrations'],
  ['SMS enabled', true, 'Text phones that gave a mobile number (needs 📣 Find Me → 📱 Set up SMS)'],
  ['SMS sender name', '', 'Your approved Semaphore sender name. Leave blank to use Semaphore\'s default.'],
  ['SMS on re-alerts', true, 'Also text again when a seat is re-alerted'],
  ['SMS confirmation text', true, 'Send a short text when someone saves their number, so typos are caught early'],
  ['SMS priority (faster)', true, 'Semaphore priority route: skips the queue, costs 2 credits per text']
];
var CHECKBOX_SETTINGS = ['Registration open', 'SMS enabled', 'SMS on re-alerts', 'SMS confirmation text', 'SMS priority (faster)'];

/* ======================= MENU ======================= */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📣 Find Me')
    .addItem('📣 Alert selected seats', 'menuAlertSelected')
    .addItem('↺ Cancel alert on selected seats', 'menuCancelSelected')
    .addItem('✖ Unregister selected seats (phone is told to register again)', 'menuUnregisterSelected')
    .addItem('⚠ Dismiss "this is my seat" claim on selected seats', 'menuDismissClaim')
    .addSeparator()
    .addItem('✓ Mark selected seats DONE (award received)', 'menuMarkDone')
    .addItem('↺ Undo DONE on selected seats (to call them again)', 'menuUndoDone')
    .addSeparator()
    .addItem('📊 Status summary', 'menuSummary')
    .addItem('↺ Cancel ALL alerts', 'menuCancelAll')
    .addItem('🗑 Clear ALL registrations (after the event)', 'menuClearAll')
    .addSeparator()
    .addItem('🔧 1. Set up / rebuild registry', 'setup')
    .addItem('🔗 2. Connect to Vercel', 'menuConnect')
    .addItem('🔑 Show shared secret', 'menuShowSecret')
    .addItem('🧪 Test connection', 'menuTestConnection')
    .addItem('📱 Set up SMS (Semaphore)', 'menuSetupSms')
    .addItem('📱 Send a test SMS', 'menuTestSms')
    .addToUi();
}

/* ======================= SETUP ======================= */
function setup() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();

  ensureSettingsSheet_();
  var settings = getSettings_();
  var map = ss.getSheetByName(settings.mapTab);
  if (!map) {
    ui.alert('Find Me', 'Seat map tab "' + settings.mapTab + '" was not found. Check the name in the "' + SET_SHEET + '" tab.', ui.ButtonSet.OK);
    return;
  }

  // 0) Existing registrations (kept) + the name cells we wrote before (not seat labels).
  var sh = ss.getSheetByName(REG_SHEET);
  var old = {}, oldNameCells = {};
  if (sh) {
    var last = sh.getLastRow(), lastCol = Math.max(sh.getLastColumn(), 1);
    if (last > 1) {
      // Map old columns to new ones by header text, so upgrades keep every registration.
      var oldHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
      var fromCol = HEADERS.map(function (h) { return oldHeaders.indexOf(h); });
      sh.getRange(2, 1, last - 1, lastCol).getValues().forEach(function (src) {
        var row = fromCol.map(function (k) { return k >= 0 ? src[k] : ''; });
        if (row[0]) old[row[0]] = row;
        if (row[C.NAME_CELL - 1]) oldNameCells[String(row[C.NAME_CELL - 1])] = true;
      });
    }
  }

  // 1) Scan the seat map for seat labels (e.g. A1 … L20).
  var range = map.getDataRange();
  var values = range.getDisplayValues();
  var bgs = range.getBackgrounds();
  var found = {}, dupes = [], seats = [];
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var raw = String(values[r][c]).trim();
      if (!raw || raw.length > 6) continue;
      if (oldNameCells[a1_(r + 1, c + 1)]) continue;
      var seat = normSeat_(raw);
      if (!SEAT_RE.test(seat)) continue;
      var a1 = map.getRange(r + 1, c + 1).getA1Notation();
      if (found[seat]) { dupes.push(seat + ' (' + found[seat] + ' & ' + a1 + ')'); continue; }
      found[seat] = a1;
      seats.push({ seat: seat, cell: a1, bg: bgs[r][c], r: r, c: c });
    }
  }
  // Name cell = the cell directly above each seat, if it's free (or holds a name we wrote).
  var seatCells = {};
  seats.forEach(function (s) { seatCells[s.cell] = true; });
  seats.forEach(function (s) {
    s.nameCell = '';
    if (s.r < 1) return;
    var above = a1_(s.r, s.c + 1);
    var v = String(values[s.r - 1][s.c] || '').trim();
    if (seatCells[above]) return;
    if (!v || oldNameCells[above]) s.nameCell = above;
  });
  if (!seats.length) { ui.alert('Find Me', 'No seat labels (like A1, B12) were found on "' + settings.mapTab + '".', ui.ButtonSet.OK); return; }
  seats.sort(function (a, b) {
    var ra = a.seat.replace(/\d+/, ''), rb = b.seat.replace(/\d+/, '');
    if (ra !== rb) return ra.length - rb.length || (ra < rb ? -1 : 1);
    return parseInt(a.seat.replace(/\D/g, ''), 10) - parseInt(b.seat.replace(/\D/g, ''), 10);
  });

  // 2) Build / merge the registry (keeps existing registrations).
  if (!sh) sh = ss.insertSheet(REG_SHEET);
  // Clear name cells that are no longer used.
  Object.keys(oldNameCells).forEach(function (cell) {
    var stillUsed = seats.some(function (s) { return s.nameCell === cell; });
    if (!stillUsed) map.getRange(cell).setValue('');
  });
  var rows = seats.map(function (s) {
    var row = old[s.seat] ? old[s.seat].slice() : blankRow_();
    row[C.SEAT - 1] = s.seat;
    row[C.MAP_CELL - 1] = s.cell;
    row[C.NAME_CELL - 1] = s.nameCell;
    row[C.ALERT - 1] = row[C.ALERT - 1] === true;
    row[C.DONE - 1] = row[C.DONE - 1] === true;
    row[C.UNREG - 1] = false;
    if (!row[C.ORIG_BG - 1]) row[C.ORIG_BG - 1] = s.bg || '#ffffff';
    row[C.VIA - 1] = via_(row);
    return row;
  });

  sh.clear();
  if (sh.getMaxColumns() < NCOLS) sh.insertColumnsAfter(sh.getMaxColumns(), NCOLS - sh.getMaxColumns());
  sh.showColumns(1, NCOLS);
  sh.getRange(1, 1, 1, NCOLS).setValues([HEADERS]).setFontWeight('bold').setBackground('#111827').setFontColor('#ffffff');
  sh.getRange(2, 1, rows.length, NCOLS).setValues(rows);
  sh.setFrozenRows(1);
  sh.getRange(2, C.ALERT, rows.length, 1).insertCheckboxes();
  sh.getRange(2, C.DONE, rows.length, 1).insertCheckboxes();
  sh.getRange(2, C.UNREG, rows.length, 1).insertCheckboxes();
  sh.getRange(2, C.REGISTERED, rows.length, 1).setNumberFormat('h:mm:ss am/pm');
  sh.getRange(2, C.ALERTED_AT, rows.length, 1).setNumberFormat('h:mm:ss am/pm');
  sh.getRange(2, C.ACK_AT, rows.length, 1).setNumberFormat('h:mm:ss am/pm');
  sh.getRange(2, C.LAST_SENT, rows.length, 1).setNumberFormat('h:mm:ss am/pm');
  sh.getRange(2, C.NAME, rows.length, 1).setNumberFormat('@');
  sh.setColumnWidth(C.NAME, 240);
  sh.setColumnWidth(C.STATUS, 110);
  sh.setColumnWidth(C.ISSUE, 360);
  sh.setColumnWidth(C.VIA, 130);
  sh.hideColumns(C.MOBILE, NCOLS - C.MOBILE + 1);

  // Name cells on the seat map: small, wrapped text so long names fit.
  var nameCells = seats.filter(function (s) { return s.nameCell; }).map(function (s) { return s.nameCell; });
  if (nameCells.length) {
    map.getRangeList(nameCells).setFontSize(8).setWrap(true).setVerticalAlignment('bottom').setHorizontalAlignment('center');
  }

  // Status colours in the registry
  var statusRange = sh.getRange(2, C.STATUS, rows.length, 1);
  var issueRange = sh.getRange(2, C.ISSUE, rows.length, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('⚠').setBackground(COLORS.issueFill).setBold(true).setRanges([issueRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(STATUS.ALERTED).setBackground(COLORS.alertedFill).setBold(true).setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(STATUS.ON_THE_WAY).setBackground(COLORS.onWayFill).setBold(true).setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(STATUS.REGISTERED).setBackground('#DBEAFE').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(STATUS.DONE).setBackground(COLORS.doneFill).setBold(true).setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Page only').setBackground('#FEF3C7').setRanges([sh.getRange(2, C.VIA, rows.length, 1)]).build()
  ]);

  // Repaint the map for any existing registrations / issues.
  var mapSheet = map;
  rows.forEach(function (row) { if (row[C.NAME - 1] || row[C.ISSUE - 1]) paintSeat_(mapSheet, row); });

  // 3) Triggers (edit checkbox + re-alert timer).
  installTriggers_();
  invalidate_();
  CacheService.getScriptCache().remove('layout');

  var msg = 'Registry ready: ' + rows.length + ' seats found on "' + settings.mapTab + '".';
  var noName = seats.filter(function (s) { return !s.nameCell; }).map(function (s) { return s.seat; });
  msg += '\nNames will show above ' + (seats.length - noName.length) + ' seats' +
    (noName.length ? ' (no free cell above: ' + noName.slice(0, 12).join(', ') + (noName.length > 12 ? '…' : '') + ')' : '') + '.';
  if (dupes.length) msg += '\n\n⚠ Duplicate seat labels (fix these on the map, then run setup again):\n' + dupes.join('\n');
  if (!PropertiesService.getScriptProperties().getProperty('VERCEL_URL')) msg += '\n\nNext: 📣 Find Me → 🔗 2. Connect to Vercel.';
  if (!PropertiesService.getScriptProperties().getProperty('SEMAPHORE_KEY')) msg += '\n\nSMS fallback is not set up yet: 📣 Find Me → 📱 Set up SMS.';
  ui.alert('Find Me', msg, ui.ButtonSet.OK);
}

function ensureSettingsSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SET_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SET_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'Notes']]).setFontWeight('bold').setBackground('#111827').setFontColor('#ffffff');
    sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 520); sh.setColumnWidth(3, 420);
    sh.setFrozenRows(1);
  }
  // Add any settings this version introduced (existing values are never touched).
  var have = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { have[String(r[0]).trim()] = true; });
  DEFAULT_SETTINGS.forEach(function (d) {
    if (have[d[0]]) return;
    var r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, 3).setValues([d]);
    sh.getRange(r, 2).setWrap(true);
    if (CHECKBOX_SETTINGS.indexOf(d[0]) >= 0) sh.getRange(r, 2).insertCheckboxes().setValue(d[1] === true);
  });
  _settings = null;
  return sh;
}

function installTriggers_() {
  var id = SpreadsheetApp.getActive().getId();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'handleEdit' || fn === 'realertTick') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(id).onEdit().create();
  ScriptApp.newTrigger('realertTick').timeBased().everyMinutes(1).create();
}

/* ======================= SETTINGS ======================= */
var _settings = null;
function getSettings_() {
  if (_settings) return _settings;
  var map = {};
  var sh = SpreadsheetApp.getActive().getSheetByName(SET_SHEET);
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { map[String(r[0]).trim()] = r[1]; });
  }
  var num = function (v, d) { var n = Number(v); return isNaN(n) || v === '' ? d : n; };
  var bool = function (k, d) { var v = map[k]; return v === undefined || v === '' ? d : v !== false && String(v).toUpperCase() !== 'FALSE'; };
  _settings = {
    title: String(map['Alert title'] || DEFAULT_SETTINGS[0][1]).slice(0, 80),
    message: String(map['Alert message'] || DEFAULT_SETTINGS[1][1]).slice(0, 300),
    mapTab: String(map['Seat map tab'] || DEFAULT_MAP).trim(),
    realertSec: num(map['Re-alert after (seconds)'], 60),
    maxAlerts: num(map['Max alerts per seat'], 3),
    registrationOpen: bool('Registration open', true),
    smsEnabled: bool('SMS enabled', true),
    smsSender: String(map['SMS sender name'] || '').trim(),
    smsOnRealert: bool('SMS on re-alerts', true),
    smsConfirm: bool('SMS confirmation text', true),
    smsPriority: bool('SMS priority (faster)', true)
  };
  return _settings;
}

/* ======================= REGISTRY ACCESS ======================= */
function getRegistry_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(REG_SHEET);
  if (!sh) throw new Error('Registry not set up. Run 📣 Find Me → Set up first.');
  var last = sh.getLastRow();
  var rows = last > 1 ? sh.getRange(2, 1, last - 1, NCOLS).getValues() : [];
  var index = {};
  rows.forEach(function (r, i) { if (r[0]) index[String(r[0])] = i; });
  return { sheet: sh, rows: rows, index: index };
}
function writeRow_(reg, i) {
  reg.rows[i][C.VIA - 1] = via_(reg.rows[i]);
  reg.sheet.getRange(i + 2, 1, 1, NCOLS).setValues([reg.rows[i]]);
}

/** How this seat will be reached: "Push ✓" (confirmed), "Push (unconfirmed)", "SMS", or "Page only". */
function via_(row) {
  if (!row[C.NAME - 1]) return '';
  var parts = [];
  if (row[C.PUSH_OK - 1]) parts.push('Push ✓');
  else if (row[C.SUB - 1]) parts.push('Push (unconfirmed)');
  if (row[C.MOBILE - 1]) parts.push('SMS');
  return parts.length ? parts.join(' + ') : 'Page only';
}

var _map = null;
function mapSheet_() {
  if (_map) return _map;
  _map = SpreadsheetApp.getActive().getSheetByName(getSettings_().mapTab);
  return _map;
}

function a1_(row, col) {
  var s = '';
  while (col > 0) { var m = (col - 1) % 26; s = String.fromCharCode(65 + m) + s; col = Math.floor((col - 1) / 26); }
  return s + row;
}
function blankRow_() { var r = []; for (var k = 0; k < NCOLS; k++) r.push(''); return r; }
function normSeat_(v) { return String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function normName_(v) { return String(v || '').replace(/^'/, '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function safeText_(v) {
  var s = String(v || '').replace(/\s+/g, ' ').trim().slice(0, 80);
  return /^[=+\-@]/.test(s) ? "'" + s : s; // block spreadsheet formula injection
}

/* ======================= MAP PAINTING ======================= */
function paintSeat_(map, row) {
  map = map || mapSheet_();
  var cell = row[C.MAP_CELL - 1];
  if (!map || !cell) return;
  var nameCell = row[C.NAME_CELL - 1];
  if (nameCell) {
    var nm = String(row[C.NAME - 1] || '');
    var nrg = map.getRange(String(nameCell));
    if (String(nrg.getDisplayValue()) !== nm) nrg.setValue(nm ? safeText_(nm) : '');
  }
  var rg = map.getRange(cell);
  var status = row[C.STATUS - 1];
  var orig = row[C.ORIG_BG - 1] || null;
  var thick = SpreadsheetApp.BorderStyle.SOLID_THICK;
  if (status === STATUS.ALERTED) {
    rg.setBackground(COLORS.alertedFill).setBorder(true, true, true, true, false, false, COLORS.alertedBorder, thick).setFontWeight('bold');
  } else if (status === STATUS.ON_THE_WAY) {
    rg.setBackground(COLORS.onWayFill).setBorder(true, true, true, true, false, false, COLORS.onWayBorder, thick).setFontWeight('bold');
  } else if (status === STATUS.DONE) {
    rg.setBackground(COLORS.doneFill).setBorder(true, true, true, true, false, false, COLORS.doneBorder, thick).setFontWeight('bold');
  } else if (row[C.ISSUE - 1]) {
    rg.setBackground(COLORS.issueFill).setBorder(true, true, true, true, false, false, COLORS.issueBorder, thick).setFontWeight('bold');
  } else if (row[C.NAME - 1]) {
    rg.setBackground(orig).setBorder(true, true, true, true, false, false, COLORS.registeredBorder, thick).setFontWeight('bold');
  } else {
    rg.setBackground(orig).setBorder(true, true, true, true, false, false, COLORS.plainBorder, SpreadsheetApp.BorderStyle.SOLID).setFontWeight('normal');
  }
}

/* ======================= ALERTS ======================= */
function alertSeats_(seats, opts) {
  opts = opts || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  var settings = getSettings_();
  var targets = [], notRegistered = [], unknown = [], alerted = [], done = [], sms = [];
  var reg;
  try {
    reg = getRegistry_();
    var now = new Date();
    seats.forEach(function (seat) {
      var i = reg.index[seat];
      if (i === undefined) { unknown.push(seat); return; }
      var row = reg.rows[i];
      if (!row[C.NAME - 1]) {
        notRegistered.push(seat);
        if (row[C.ALERT - 1] === true) { row[C.ALERT - 1] = false; writeRow_(reg, i); }
        return;
      }
      if (row[C.DONE - 1] === true) {
        done.push(seat);
        if (row[C.ALERT - 1] === true) { row[C.ALERT - 1] = false; writeRow_(reg, i); }
        return;
      }
      var count = Number(row[C.COUNT - 1] || 0) + 1;
      var id = now.getTime() + '-' + count;
      row[C.ALERT - 1] = true;
      row[C.STATUS - 1] = STATUS.ALERTED;
      if (!opts.realert || !row[C.ALERTED_AT - 1]) row[C.ALERTED_AT - 1] = now;
      row[C.ACK_AT - 1] = '';
      row[C.COUNT - 1] = count;
      row[C.ALERT_ID - 1] = id;
      row[C.LAST_SENT - 1] = now;
      writeRow_(reg, i);
      paintSeat_(null, row);
      alerted.push(seat);
      if (row[C.SUB - 1]) targets.push({ type: 'alert', seat: seat, alertId: id, device: String(row[C.DEVICE - 1]), subscription: String(row[C.SUB - 1]) });
      if (row[C.MOBILE - 1] && (!opts.realert || settings.smsOnRealert)) sms.push({ seat: seat, number: String(row[C.MOBILE - 1]) });
    });
    invalidate_();
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  var push = { sent: [], expired: [], expiredTargets: [], failed: [], error: '' };
  if (targets.length) push = sendPush_(targets, settings);
  if (push.expiredTargets.length) clearSubscriptions_(push.expiredTargets);

  var smsOut = { sent: [], failed: [], error: '' };
  if (sms.length) {
    smsOut = sendSms_(sms.map(function (m) {
      return { seat: m.seat, number: m.number, text: smsText_(settings, m.seat) };
    }), settings);
  }
  var reached = {};
  push.sent.forEach(function (x) { reached[x] = true; });
  smsOut.sent.forEach(function (x) { reached[x] = true; });
  var pageOnly = alerted.filter(function (x) { return !reached[x]; });

  return { alerted: alerted, notRegistered: notRegistered, unknown: unknown, done: done, push: push, sms: smsOut, pageOnly: pageOnly, noPush: alerted.length - targets.length };
}

function sendPush_(targets, settings) {
  var props = PropertiesService.getScriptProperties();
  var base = props.getProperty('VERCEL_URL');
  var secret = props.getProperty('SHARED_SECRET');
  var out = { sent: [], expired: [], expiredTargets: [], failed: [], error: '' };
  if (!base || !secret) { out.error = 'Not connected to Vercel — phones with the page open will still flash.'; return out; }
  var serverTime = Date.now();
  targets.forEach(function (t) { t.serverTime = serverTime; });
  for (var i = 0; i < targets.length; i += 100) {
    var chunk = targets.slice(i, i + 100);
    try {
      var resp = UrlFetchApp.fetch(base + '/api/push', {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ secret: secret, title: settings.title, message: settings.message, targets: chunk }),
        muteHttpExceptions: true
      });
      var data = JSON.parse(resp.getContentText() || '{}');
      if (resp.getResponseCode() !== 200 || !data.ok) {
        out.error = 'Push service error (' + resp.getResponseCode() + '): ' + (data.error || '');
        chunk.forEach(function (t) { out.failed.push(t.seat); });
        continue;
      }
      out.sent = out.sent.concat(data.sent || []);
      out.expired = out.expired.concat(data.expired || []);
      (data.expiredIdx || []).forEach(function (k) { if (chunk[k]) out.expiredTargets.push(chunk[k]); });
      out.failed = out.failed.concat(data.failed || []);
    } catch (e) {
      out.error = 'Could not reach Vercel: ' + e.message;
      chunk.forEach(function (t) { out.failed.push(t.seat); });
    }
  }
  return out;
}

/** Removes push subscriptions the push service says are gone (only if the row still holds that exact one). */
function clearSubscriptions_(expiredTargets) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var reg = getRegistry_();
    expiredTargets.forEach(function (t) {
      var i = reg.index[t.seat];
      if (i === undefined) return;
      if (String(reg.rows[i][C.SUB - 1]) === String(t.subscription)) { reg.rows[i][C.SUB - 1] = ''; writeRow_(reg, i); }
    });
  } finally { lock.releaseLock(); }
}

/* ======================= ISSUES (claims + duplicate names) ======================= */
function parseClaim_(v) {
  if (!v) return null;
  try { var c = JSON.parse(String(v)); return c && c.device ? c : null; } catch (e) { return null; }
}
function fmtTime_(ms) {
  var tz = (typeof Session !== 'undefined' && Session.getScriptTimeZone()) || 'Asia/Manila';
  return Utilities.formatDate(new Date(ms), tz, 'h:mm a');
}

/** Recomputes the ⚠ Issue column for every seat; writes + repaints only rows that changed. */
function refreshIssues_(reg) {
  var byName = {};
  reg.rows.forEach(function (r) {
    var n = normName_(r[C.NAME - 1]);
    if (n) (byName[n] = byName[n] || []).push(String(r[0]));
  });
  reg.rows.forEach(function (r, i) {
    var parts = [];
    var claim = parseClaim_(r[C.CLAIM - 1]);
    if (claim) parts.push('⚠ ' + claim.name + ' says this is their seat (' + fmtTime_(claim.at) + '). Check, then unregister the wrong person.');
    var n = normName_(r[C.NAME - 1]);
    if (n && byName[n].length > 1) {
      var others = byName[n].filter(function (s) { return s !== String(r[0]); });
      parts.push('⚠ Same name also registered at ' + others.join(', '));
    }
    var text = parts.join('  |  ');
    if (text !== String(r[C.ISSUE - 1] || '')) {
      r[C.ISSUE - 1] = text;
      writeRow_(reg, i);
      paintSeat_(null, r);
    }
  });
}

function noticeTarget_(kind, seat, subscription, device) {
  var n = NOTICE[kind];
  return {
    type: kind, seat: seat, alertId: kind + '-' + Date.now(), subscription: String(subscription), device: device || '',
    title: n.title, message: n.message.replace('{seat}', seat)
  };
}

function cancelSeats_(seats) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var n = 0;
  try {
    var reg = getRegistry_();
    seats.forEach(function (seat) {
      var i = reg.index[seat];
      if (i === undefined) return;
      var row = reg.rows[i];
      row[C.ALERT - 1] = false;
      row[C.STATUS - 1] = !row[C.NAME - 1] ? STATUS.NONE : (row[C.DONE - 1] === true ? STATUS.DONE : STATUS.REGISTERED);
      row[C.ALERT_ID - 1] = '';
      row[C.COUNT - 1] = '';
      row[C.ALERTED_AT - 1] = '';
      row[C.ACK_AT - 1] = '';
      row[C.LAST_SENT - 1] = '';
      writeRow_(reg, i);
      paintSeat_(null, row);
      n++;
    });
    invalidate_();
  } finally { lock.releaseLock(); }
  return n;
}

function clearRow_(row) {
  [C.NAME, C.REGISTERED, C.STATUS, C.ALERTED_AT, C.ACK_AT, C.COUNT, C.DEVICE, C.SUB, C.ALERT_ID, C.LAST_SENT, C.MOBILE, C.PUSH_OK]
    .forEach(function (col) { row[col - 1] = ''; });
  row[C.ALERT - 1] = false;
  row[C.DONE - 1] = false;
  row[C.UNREG - 1] = false;
}

/**
 * Frees seat row i. If someone said "this is my seat", they get it automatically.
 * Returns push targets to send (register-again notice and/or seat-granted notice).
 */
function releaseSeat_(reg, i, notify) {
  var row = reg.rows[i];
  var seat = String(row[0]);
  var targets = [];
  if (notify && row[C.NAME - 1] && row[C.SUB - 1]) targets.push(noticeTarget_('unregistered', seat, row[C.SUB - 1]));
  var claim = parseClaim_(row[C.CLAIM - 1]);
  clearRow_(row);
  if (claim) {
    // The claimant may hold another seat by mistake — free it (no notice to themselves).
    reg.rows.forEach(function (r, j) {
      if (j !== i && r[C.DEVICE - 1] === claim.device) { clearRow_(r); writeRow_(reg, j); paintSeat_(null, r); }
    });
    row[C.NAME - 1] = safeText_(claim.name);
    row[C.REGISTERED - 1] = new Date();
    row[C.DEVICE - 1] = claim.device;
    row[C.SUB - 1] = claim.sub || '';
    row[C.STATUS - 1] = STATUS.REGISTERED;
    if (claim.sub) targets.push(noticeTarget_('granted', seat, claim.sub, claim.device));
  }
  row[C.CLAIM - 1] = '';
  writeRow_(reg, i);
  paintSeat_(null, row);
  return targets;
}

function unregisterSeats_(seats, opts) {
  opts = opts || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var n = 0, granted = [], targets = [];
  try {
    var reg = getRegistry_();
    seats.forEach(function (seat) {
      var i = reg.index[seat];
      if (i === undefined) return;
      var hadClaim = !!parseClaim_(reg.rows[i][C.CLAIM - 1]);
      if (opts.wipe) { clearRow_(reg.rows[i]); reg.rows[i][C.CLAIM - 1] = ''; writeRow_(reg, i); paintSeat_(null, reg.rows[i]); }
      else {
        targets = targets.concat(releaseSeat_(reg, i, opts.notify !== false));
        if (hadClaim) granted.push(seat + ' → ' + reg.rows[i][C.NAME - 1]);
      }
      n++;
    });
    refreshIssues_(reg);
    invalidate_();
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  var push = targets.length ? sendPush_(targets, getSettings_()) : { sent: [], expiredTargets: [], failed: [], error: '' };
  if (push.expiredTargets.length) clearSubscriptions_(push.expiredTargets);
  var unreg = targets.filter(function (t) { return t.type === 'unregistered'; }).length;
  return { count: n, granted: granted, notified: push.error ? 0 : unreg, pushError: push.error };
}

function dismissClaims_(seats) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var targets = [], n = 0;
  try {
    var reg = getRegistry_();
    seats.forEach(function (seat) {
      var i = reg.index[seat];
      if (i === undefined) return;
      var claim = parseClaim_(reg.rows[i][C.CLAIM - 1]);
      if (!claim) return;
      if (claim.sub) targets.push(noticeTarget_('dismissed', seat, claim.sub));
      reg.rows[i][C.CLAIM - 1] = '';
      writeRow_(reg, i);
      n++;
    });
    refreshIssues_(reg);
    invalidate_();
  } finally { lock.releaseLock(); }
  if (targets.length) sendPush_(targets, getSettings_());
  return n;
}

/** Runs every minute: re-send to seats that are still "Alerted". */
function realertTick() {
  var s = getSettings_();
  if (!s.realertSec || s.realertSec <= 0) return;
  var reg;
  try { reg = getRegistry_(); } catch (e) { return; }
  var now = Date.now(), due = [];
  reg.rows.forEach(function (row) {
    if (row[C.STATUS - 1] !== STATUS.ALERTED) return;
    if (Number(row[C.COUNT - 1] || 0) >= s.maxAlerts) return;
    var last = row[C.LAST_SENT - 1] instanceof Date ? row[C.LAST_SENT - 1].getTime() : 0;
    if (now - last >= (s.realertSec - 5) * 1000) due.push(String(row[C.SEAT - 1]));
  });
  if (due.length) alertSeats_(due, { realert: true });
}

/* ======================= SELECTION HELPERS ======================= */
function selectedSeats_() {
  var sh = SpreadsheetApp.getActiveSheet();
  var list = SpreadsheetApp.getActiveRangeList();
  if (!list) return [];
  var seats = {};
  var isReg = sh.getName() === REG_SHEET;
  list.getRanges().forEach(function (rg) {
    if (isReg) {
      var startRow = Math.max(rg.getRow(), 2);
      var endRow = rg.getLastRow();
      if (endRow < startRow) return;
      sh.getRange(startRow, C.SEAT, endRow - startRow + 1, 1).getValues().forEach(function (r) {
        var s = normSeat_(r[0]); if (SEAT_RE.test(s)) seats[s] = true;
      });
    } else {
      var r0 = rg.getRow(), c0 = rg.getColumn();
      rg.getDisplayValues().forEach(function (r, i) {
        r.forEach(function (v, j) {
          var byName = nameCellMap()[a1_(r0 + i, c0 + j)];
          if (byName) { seats[byName] = true; return; }
          var raw = String(v).trim();
          if (!raw || raw.length > 6) return;
          var s = normSeat_(raw); if (SEAT_RE.test(s)) seats[s] = true;
        });
      });
    }
  });
  return Object.keys(seats);

  function nameCellMap() {
    if (nameCellMap.cache) return nameCellMap.cache;
    var m = {};
    try {
      getRegistry_().rows.forEach(function (row) { if (row[C.NAME_CELL - 1]) m[String(row[C.NAME_CELL - 1])] = String(row[0]); });
    } catch (e) { /* not set up yet */ }
    return (nameCellMap.cache = m);
  }
}

/* ======================= MENU ACTIONS ======================= */
function menuAlertSelected() {
  var seats = selectedSeats_();
  var ss = SpreadsheetApp.getActive();
  if (!seats.length) { ss.toast('Select one or more seat cells on the seat map (or rows in the registry) first.', 'Find Me', 6); return; }
  var r = alertSeats_(seats);
  ss.toast(summaryText_(r), '📣 Find Me', 10);
}

function summaryText_(r) {
  var parts = [];
  if (r.alerted.length) parts.push('Alerted ' + r.alerted.length + ': ' + r.alerted.join(', '));
  if (r.push.sent.length) parts.push('Push sent to ' + r.push.sent.length);
  if (r.sms && r.sms.sent.length) parts.push('SMS sent to ' + r.sms.sent.length);
  if (r.pageOnly && r.pageOnly.length) parts.push('⚠ Only reachable if their page is open — watch these: ' + r.pageOnly.join(', '));
  if (r.done && r.done.length) parts.push('Skipped (marked DONE — untick DONE to alert again): ' + r.done.join(', '));
  if (r.sms && r.sms.failed.length) parts.push('⚠ SMS failed: ' + r.sms.failed.join(', '));
  if (r.sms && r.sms.error) parts.push('⚠ SMS: ' + r.sms.error);
  if (r.notRegistered.length) parts.push('⚠ Not registered — send an usher: ' + r.notRegistered.join(', '));
  if (r.unknown.length) parts.push('⚠ Not on map: ' + r.unknown.join(', '));
  if (r.push.expired.length) parts.push('⚠ Notifications expired: ' + r.push.expired.join(', '));
  if (r.push.failed.length) parts.push('⚠ Push failed: ' + r.push.failed.join(', '));
  if (r.push.error) parts.push('⚠ ' + r.push.error);
  return parts.join('\n') || 'Nothing to alert.';
}

function menuCancelSelected() {
  var seats = selectedSeats_();
  var n = seats.length ? cancelSeats_(seats) : 0;
  SpreadsheetApp.getActive().toast(n ? 'Cancelled alert on ' + n + ' seat(s).' : 'Select seats first.', 'Find Me', 5);
}

function menuUnregisterSelected() {
  var ui = SpreadsheetApp.getUi();
  var seats = selectedSeats_();
  if (!seats.length) { SpreadsheetApp.getActive().toast('Select seats first.', 'Find Me', 5); return; }
  if (ui.alert('Unregister ' + seats.length + ' seat(s)?', seats.join(', ') +
    '\n\n• The name and phone link are removed.\n• That phone gets a notification to register again.' +
    '\n• If someone said "this is my seat" (red seat), they get the seat automatically.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var r = unregisterSeats_(seats, { notify: true });
  SpreadsheetApp.getActive().toast(unregText_(r, seats), 'Find Me', 8);
}

function unregText_(r, seats) {
  var msg = 'Unregistered ' + r.count + ' seat(s): ' + seats.join(', ');
  if (r.notified) msg += '\nRegister-again notice sent to ' + r.notified + ' phone(s).';
  if (r.granted.length) msg += '\nSeat given to: ' + r.granted.join(', ');
  if (r.pushError) msg += '\n⚠ ' + r.pushError;
  return msg;
}

/* ======================= DONE (award received) ======================= */
function markDone_(seats) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var n = 0;
  try {
    var reg = getRegistry_();
    seats.forEach(function (seat) {
      var i = reg.index[seat];
      if (i === undefined) return;
      var row = reg.rows[i];
      if (!row[C.NAME - 1]) { if (row[C.DONE - 1] === true) { row[C.DONE - 1] = false; writeRow_(reg, i); } return; }
      row[C.DONE - 1] = true;
      row[C.ALERT - 1] = false;
      row[C.STATUS - 1] = STATUS.DONE;   // also stops any running alert / re-alerts
      writeRow_(reg, i);
      paintSeat_(null, row);
      n++;
    });
    invalidate_();
  } finally { lock.releaseLock(); }
  return n;
}

function undoDone_(seats) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var n = 0;
  try {
    var reg = getRegistry_();
    seats.forEach(function (seat) {
      var i = reg.index[seat];
      if (i === undefined) return;
      var row = reg.rows[i];
      row[C.DONE - 1] = false;
      row[C.ALERT - 1] = false;
      row[C.STATUS - 1] = row[C.NAME - 1] ? STATUS.REGISTERED : STATUS.NONE;
      // Fresh start for the next call (re-alert counter etc.).
      [C.ALERTED_AT, C.ACK_AT, C.COUNT, C.ALERT_ID, C.LAST_SENT].forEach(function (col) { row[col - 1] = ''; });
      writeRow_(reg, i);
      paintSeat_(null, row);
      n++;
    });
    invalidate_();
  } finally { lock.releaseLock(); }
  return n;
}

function menuMarkDone() {
  var seats = selectedSeats_();
  if (!seats.length) { SpreadsheetApp.getActive().toast('Select the seat(s) first.', 'Find Me', 5); return; }
  var n = markDone_(seats);
  SpreadsheetApp.getActive().toast(n ? 'Marked DONE: ' + n + ' seat(s). Use "Undo DONE" to call them again.' : 'No registered seats selected.', 'Find Me', 6);
}

function menuUndoDone() {
  var seats = selectedSeats_();
  if (!seats.length) { SpreadsheetApp.getActive().toast('Select the seat(s) first.', 'Find Me', 5); return; }
  var n = undoDone_(seats);
  SpreadsheetApp.getActive().toast('Ready to alert again: ' + n + ' seat(s).', 'Find Me', 5);
}

function menuDismissClaim() {
  var seats = selectedSeats_();
  if (!seats.length) { SpreadsheetApp.getActive().toast('Select the red seat(s) first.', 'Find Me', 5); return; }
  var n = dismissClaims_(seats);
  SpreadsheetApp.getActive().toast(n ? 'Dismissed ' + n + ' claim(s). Those phones were asked to re-check their seat.' : 'No claims on the selected seats.', 'Find Me', 6);
}

function menuSummary() {
  var reg = getRegistry_();
  var c = { total: reg.rows.length, registered: 0, push: 0, pushOk: 0, sms: 0, pageOnly: 0, alerted: 0, onWay: 0, done: 0 };
  var waiting = [], issues = [], pageOnly = [];
  reg.rows.forEach(function (r) {
    if (r[C.NAME - 1]) c.registered++;
    if (r[C.SUB - 1]) c.push++;
    if (r[C.NAME - 1] && r[C.PUSH_OK - 1]) c.pushOk++;
    if (r[C.NAME - 1] && r[C.MOBILE - 1]) c.sms++;
    if (r[C.NAME - 1] && !r[C.PUSH_OK - 1] && !r[C.MOBILE - 1]) { c.pageOnly++; pageOnly.push(r[0]); }
    if (r[C.STATUS - 1] === STATUS.DONE) c.done++;
    if (r[C.STATUS - 1] === STATUS.ALERTED) { c.alerted++; waiting.push(r[0]); }
    if (r[C.STATUS - 1] === STATUS.ON_THE_WAY) c.onWay++;
    if (r[C.ISSUE - 1]) issues.push(r[0]);
  });
  SpreadsheetApp.getUi().alert('📊 Find Me status',
    'Seats: ' + c.total + '\nRegistered: ' + c.registered +
    '\n\nReachable by push (confirmed): ' + c.pushOk + '\nReachable by SMS: ' + c.sms +
    '\n⚠ Not confirmed (page-open only or push unconfirmed): ' + c.pageOnly + (pageOnly.length ? ' (' + pageOnly.slice(0, 30).join(', ') + (pageOnly.length > 30 ? '…' : '') + ')' : '') +
    '\n\nAlerted, waiting for reply: ' + c.alerted + (waiting.length ? ' (' + waiting.join(', ') + ')' : '') +
    '\nOn the way: ' + c.onWay + '\nDone (award received): ' + c.done +
    '\n\n⚠ Seats needing staff (red): ' + issues.length + (issues.length ? ' (' + issues.join(', ') + ')' : ''), SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuCancelAll() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('Cancel ALL alerts?', 'Every seat goes back to "Registered". Names stay.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var reg = getRegistry_();
  var seats = reg.rows.filter(function (r) { return r[C.STATUS - 1] === STATUS.ALERTED || r[C.STATUS - 1] === STATUS.ON_THE_WAY; })
    .map(function (r) { return String(r[0]); });
  var n = cancelSeats_(seats);
  SpreadsheetApp.getActive().toast('Reset ' + n + ' seat(s).', 'Find Me', 5);
}

function menuClearAll() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('Clear ALL registrations?', 'Deletes every name, mobile number and phone link and restores the seat map colours.\nUse this after the event (Data Privacy). This cannot be undone.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var reg = getRegistry_();
  var r = unregisterSeats_(reg.rows.map(function (row) { return String(row[0]); }), { wipe: true });
  SpreadsheetApp.getActive().toast('Cleared ' + r.count + ' seats.', 'Find Me', 5);
}

function menuConnect() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var cur = props.getProperty('VERCEL_URL') || '';
  var res = ui.prompt('🔗 Connect to Vercel', 'Paste your Vercel site address (e.g. https://find-me.vercel.app)' + (cur ? '\nCurrent: ' + cur : ''), ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var url = res.getResponseText().trim().replace(/\/+$/, '');
  if (!/^https:\/\//.test(url)) { ui.alert('Please paste the full address starting with https://'); return; }
  props.setProperty('VERCEL_URL', url);
  if (!props.getProperty('SHARED_SECRET')) props.setProperty('SHARED_SECRET', Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''));
  ui.alert('Connected',
    'Saved: ' + url + '\n\nIn Vercel → Settings → Environment Variables, set:\n\nSHARED_SECRET =\n' + props.getProperty('SHARED_SECRET') +
    '\n\nThen redeploy Vercel and run 📣 Find Me → 🧪 Test connection.', ui.ButtonSet.OK);
}

function menuShowSecret() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SHARED_SECRET')) props.setProperty('SHARED_SECRET', Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''));
  SpreadsheetApp.getUi().alert('SHARED_SECRET (copy into Vercel)', props.getProperty('SHARED_SECRET'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuTestConnection() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var base = props.getProperty('VERCEL_URL');
  if (!base) { ui.alert('Not connected yet. Use 📣 Find Me → 🔗 Connect to Vercel.'); return; }
  var lines = [];
  try {
    var cfg = UrlFetchApp.fetch(base + '/api/config', { muteHttpExceptions: true });
    var c = JSON.parse(cfg.getContentText() || '{}');
    lines.push((cfg.getResponseCode() === 200 ? '✅' : '❌') + ' Site reachable (' + cfg.getResponseCode() + ')');
    lines.push((c.vapidPublicKey ? '✅' : '❌') + ' Push keys set in Vercel');
  } catch (e) { lines.push('❌ Site not reachable: ' + e.message); }
  try {
    var p = UrlFetchApp.fetch(base + '/api/push', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ secret: props.getProperty('SHARED_SECRET'), targets: [] })
    });
    lines.push((p.getResponseCode() === 200 ? '✅ Shared secret matches' : '❌ Push endpoint: HTTP ' + p.getResponseCode() + ' (check SHARED_SECRET in Vercel)'));
  } catch (e) { lines.push('❌ Push endpoint error: ' + e.message); }
  try {
    var a = UrlFetchApp.fetch(base + '/api/alerts', { muteHttpExceptions: true });
    lines.push((a.getResponseCode() === 200 ? '✅' : '❌') + ' Vercel can read this sheet' + (a.getResponseCode() === 200 ? '' : ' (check APPS_SCRIPT_URL in Vercel and that the web app is deployed for "Anyone")'));
  } catch (e) { lines.push('❌ Alerts endpoint error: ' + e.message); }
  ui.alert('🧪 Connection test', lines.join('\n'), ui.ButtonSet.OK);
}

/* ======================= EDIT TRIGGER ======================= */
/** Installable onEdit (created by setup). Handles the ALERT checkboxes + cache refresh. */
function handleEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  var name = sh.getName();
  if (name === SET_SHEET) { _settings = null; invalidate_(); return; }
  if (name !== REG_SHEET) return;
  invalidate_();
  var rg = e.range;
  if (rg.getColumn() <= C.NAME && rg.getLastColumn() >= C.NAME) {
    var lk = LockService.getScriptLock(); lk.waitLock(20000);
    try { refreshIssues_(getRegistry_()); } finally { lk.releaseLock(); }
  }
  var touches = function (col) { return rg.getColumn() <= col && rg.getLastColumn() >= col; };
  if (!touches(C.ALERT) && !touches(C.DONE) && !touches(C.UNREG)) return;
  var startRow = Math.max(rg.getRow(), 2);
  var n = rg.getLastRow() - startRow + 1;
  if (n <= 0) return;
  var vals = sh.getRange(startRow, 1, n, NCOLS).getValues();
  var toAlert = [], toCancel = [], toDone = [], toUndo = [], toUnreg = [], emptyUnreg = [];
  vals.forEach(function (r, k) {
    var seat = normSeat_(r[C.SEAT - 1]);
    if (!SEAT_RE.test(seat)) return;
    if (touches(C.UNREG) && r[C.UNREG - 1] === true) {
      if (r[C.NAME - 1]) toUnreg.push(seat); else emptyUnreg.push(startRow + k);
      return;
    }
    if (touches(C.DONE)) {
      if (r[C.DONE - 1] === true && r[C.STATUS - 1] !== STATUS.DONE) toDone.push(seat);
      if (r[C.DONE - 1] !== true && r[C.STATUS - 1] === STATUS.DONE) toUndo.push(seat);
    }
    if (touches(C.ALERT)) {
      if (r[C.ALERT - 1] === true && r[C.STATUS - 1] !== STATUS.ALERTED) toAlert.push(seat);
      if (r[C.ALERT - 1] !== true && r[C.STATUS - 1] === STATUS.ALERTED) toCancel.push(seat);
    }
  });
  var ss = SpreadsheetApp.getActive();
  emptyUnreg.forEach(function (rowNum) { sh.getRange(rowNum, C.UNREG).setValue(false); });
  if (toUnreg.length) {
    var u = unregisterSeats_(toUnreg, { notify: true });
    ss.toast(unregText_(u, toUnreg), 'Find Me', 8);
  }
  if (toDone.length) { markDone_(toDone); ss.toast('Marked DONE: ' + toDone.join(', ') + '\nUntick DONE to call them again.', 'Find Me', 6); }
  if (toUndo.length) { undoDone_(toUndo); ss.toast('Ready to alert again: ' + toUndo.join(', '), 'Find Me', 6); }
  if (toAlert.length) ss.toast(summaryText_(alertSeats_(toAlert)), '📣 Find Me', 8);
  if (toCancel.length) cancelSeats_(toCancel);
}

/* ======================= WEB APP (called by Vercel) ======================= */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function checkSecret_(s) {
  var secret = PropertiesService.getScriptProperties().getProperty('SHARED_SECRET');
  return !!secret && s === secret;
}
function invalidate_() { CacheService.getScriptCache().remove('alerts'); }

function alertsPayload_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('alerts');
  if (hit) return JSON.parse(hit);
  var startedAt = Date.now(); // stamped BEFORE reading, so a read that races a change is never newer than it
  var s = getSettings_();
  var reg = getRegistry_();
  var alerts = [];
  reg.rows.forEach(function (r) {
    if (r[C.STATUS - 1] === STATUS.ALERTED && r[C.ALERT_ID - 1]) alerts.push({ seat: String(r[0]), id: String(r[C.ALERT_ID - 1]) });
  });
  // Anonymous phone code → seat, so phones learn about staff changes (no names are published).
  var devices = {}, claims = {};
  reg.rows.forEach(function (r) {
    if (r[C.DEVICE - 1] && r[C.NAME - 1]) devices[devHash_(r[C.DEVICE - 1])] = String(r[0]);
    var cl = parseClaim_(r[C.CLAIM - 1]);
    if (cl) claims[devHash_(cl.device)] = String(r[0]);
  });
  var payload = {
    ok: true, title: s.title, message: s.message, registrationOpen: s.registrationOpen,
    smsAvailable: !!(s.smsEnabled && PropertiesService.getScriptProperties().getProperty('SEMAPHORE_KEY')),
    alerts: alerts, devices: devices, claims: claims, generatedAt: startedAt
  };
  cache.put('alerts', JSON.stringify(payload), 10);
  return payload;
}

function devHash_(device) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(device), Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < 8; i++) { var b = (bytes[i] + 256) % 256; hex += (b < 16 ? '0' : '') + b.toString(16); }
  return hex;
}

/** Seat positions for the hall map on the phone's confirm step. */
function layoutPayload_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('layout');
  if (hit) return JSON.parse(hit);
  var reg = getRegistry_();
  var seats = [];
  reg.rows.forEach(function (r) {
    var m = String(r[C.MAP_CELL - 1]).match(/^([A-Z]+)(\d+)$/);
    if (!m) return;
    var col = 0;
    for (var k = 0; k < m[1].length; k++) col = col * 26 + (m[1].charCodeAt(k) - 64);
    seats.push([String(r[0]), Number(m[2]), col]);
  });
  var stage = null;
  var map = mapSheet_();
  if (map) {
    var vals = map.getDataRange().getDisplayValues();
    for (var i = 0; i < vals.length && !stage; i++) {
      for (var j = 0; j < vals[i].length; j++) {
        if (/^\s*stage\s*$/i.test(String(vals[i][j]))) { stage = [i + 1, j + 1]; break; }
      }
    }
  }
  var payload = { ok: true, seats: seats, stage: stage };
  cache.put('layout', JSON.stringify(payload), 600);
  return payload;
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!checkSecret_(p.secret)) return json_({ ok: false, error: 'unauthorized' });
  try {
    if (p.action === 'alerts') return json_(alertsPayload_());
    if (p.action === 'layout') return json_(layoutPayload_());
    return json_({ ok: true, service: 'find-me' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  var d;
  try { d = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ ok: false, code: 'BAD_REQUEST' }); }
  if (!checkSecret_(d.secret)) return json_({ ok: false, error: 'unauthorized' });
  try {
    if (d.action === 'register') return json_(register_(d));
    if (d.action === 'ack') return json_(ack_(d));
    if (d.action === 'claim') return json_(claim_(d));
    if (d.action === 'unclaim') return json_(unclaim_(d));
    if (d.action === 'mobile') return json_(setMobile_(d));
    if (d.action === 'pushok') return json_(pushOk_(d));
    return json_({ ok: false, code: 'UNKNOWN_ACTION' });
  } catch (err) {
    return json_({ ok: false, code: 'ERROR', error: String(err.message || err) });
  }
}

function register_(d) {
  var s = getSettings_();
  if (!s.registrationOpen) return { ok: false, code: 'CLOSED' };
  var seat = normSeat_(d.seat);
  var name = String(d.name || '').replace(/\s+/g, ' ').trim().slice(0, 80);
  var device = String(d.deviceId || '');
  if (name.length < 2) return { ok: false, code: 'NAME_REQUIRED' };
  if (!SEAT_RE.test(seat)) return { ok: false, code: 'SEAT_INVALID' };
  if (!device) return { ok: false, code: 'BAD_REQUEST' };

  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  var result, pending = [];
  try {
    var reg = getRegistry_();
    var i = reg.index[seat];
    if (i === undefined) return { ok: false, code: 'SEAT_NOT_FOUND' };
    var row = reg.rows[i];
    var existing = row[C.NAME - 1];
    if (existing && row[C.DEVICE - 1] !== device && normName_(existing) !== normName_(name)) {
      return { ok: false, code: 'SEAT_TAKEN' };
    }
    pending = writeRegistration_(reg, i, name, device, d.subscription);
    refreshIssues_(reg);
    invalidate_();
    result = {
      ok: true, seat: seat, status: row[C.STATUS - 1], serverTime: Date.now(),
      alertId: row[C.STATUS - 1] === STATUS.ALERTED ? String(row[C.ALERT_ID - 1]) : ''
    };
  } finally {
    lock.releaseLock();
  }
  if (pending.length) sendPush_(pending, getSettings_());
  return result;
}

/** Puts name/device on row i. Frees any other seat or claim this phone held. Returns push targets. */
function writeRegistration_(reg, i, name, device, subscription) {
  var targets = [], carriedMobile = '';
  reg.rows.forEach(function (r, j) {
    if (j === i) return;
    if (r[C.DEVICE - 1] === device) {                                                       // phone moved seats
      if (r[C.MOBILE - 1]) carriedMobile = r[C.MOBILE - 1];
      targets = targets.concat(releaseSeat_(reg, j, false));
    }
    var cl = parseClaim_(r[C.CLAIM - 1]);
    if (cl && cl.device === device) { r[C.CLAIM - 1] = ''; writeRow_(reg, j); }       // drop old claim
  });
  var row = reg.rows[i];
  var cl2 = parseClaim_(row[C.CLAIM - 1]);
  if (cl2 && cl2.device === device) row[C.CLAIM - 1] = '';
  row[C.NAME - 1] = safeText_(name);
  if (!row[C.REGISTERED - 1]) row[C.REGISTERED - 1] = new Date();
  var newSub = subscription ? String(subscription).slice(0, 2000) : '';
  if (String(row[C.SUB - 1] || '') !== newSub || row[C.DEVICE - 1] !== device) row[C.PUSH_OK - 1] = '';
  row[C.DEVICE - 1] = device;
  row[C.SUB - 1] = newSub;
  if (carriedMobile && !row[C.MOBILE - 1]) row[C.MOBILE - 1] = carriedMobile;
  if (!row[C.STATUS - 1]) row[C.STATUS - 1] = STATUS.REGISTERED;
  writeRow_(reg, i);
  paintSeat_(null, row);
  return targets;
}

/* ======================= PUSH CONFIRMATION + SMS ======================= */
/** The phone's service worker received a push → this phone is confirmed reachable. */
function pushOk_(d) {
  var seat = normSeat_(d.seat);
  var device = String(d.deviceId || '');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var reg = getRegistry_();
    var i = reg.index[seat];
    if (i === undefined || reg.rows[i][C.DEVICE - 1] !== device) return { ok: false, code: 'NOT_YOUR_SEAT' };
    if (!reg.rows[i][C.PUSH_OK - 1]) { reg.rows[i][C.PUSH_OK - 1] = new Date(); writeRow_(reg, i); }
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function normMobile_(v) {
  var d = String(v || '').replace(/[^0-9]/g, '');
  if (/^639\d{9}$/.test(d)) d = '0' + d.slice(2);
  if (/^9\d{9}$/.test(d)) d = '0' + d;
  return /^09\d{9}$/.test(d) ? d : '';
}

/** Phone saved (or removed) a mobile number for SMS alerts. */
function setMobile_(d) {
  var seat = normSeat_(d.seat);
  var device = String(d.deviceId || '');
  var remove = d.mobile === '' || d.mobile === null;
  var mobile = remove ? '' : normMobile_(d.mobile);
  if (!remove && !mobile) return { ok: false, code: 'MOBILE_INVALID' };
  var s = getSettings_();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var changed = false;
  try {
    var reg = getRegistry_();
    var i = reg.index[seat];
    if (i === undefined || reg.rows[i][C.DEVICE - 1] !== device) return { ok: false, code: 'NOT_YOUR_SEAT' };
    changed = String(reg.rows[i][C.MOBILE - 1] || '') !== mobile;
    reg.rows[i][C.MOBILE - 1] = mobile;
    writeRow_(reg, i);
  } finally { lock.releaseLock(); }
  var smsReady = s.smsEnabled && !!PropertiesService.getScriptProperties().getProperty('SEMAPHORE_KEY');
  var confirm = { sent: [], failed: [], error: '' };
  if (mobile && changed && smsReady && s.smsConfirm) {
    confirm = sendSms_([{ seat: seat, number: mobile, text: 'Find Me: You will get a text here when it is your turn (Seat ' + seat + '). Please stay seated.' }], s);
  }
  return { ok: true, mobile: mobile, smsReady: smsReady, confirmSent: confirm.sent.length > 0 };
}

function smsText_(s, seat) {
  var t = s.title + ' (Seat ' + seat + '). ' + s.message;
  if (/^test/i.test(t)) t = 'Find Me: ' + t; // Semaphore drops messages that start with "TEST"
  return t.slice(0, 300);
}

/** Sends texts via Semaphore. items: [{ seat, number, text }] */
function sendSms_(items, s) {
  var out = { sent: [], failed: [], error: '' };
  s = s || getSettings_();
  if (!s.smsEnabled) { out.error = 'SMS is turned off in Find Me Settings.'; return out; }
  var key = PropertiesService.getScriptProperties().getProperty('SEMAPHORE_KEY');
  if (!key) { out.error = 'SMS not set up — 📣 Find Me → 📱 Set up SMS.'; items.forEach(function (m) { out.failed.push(m.seat); }); return out; }
  var url = 'https://api.semaphore.co/api/v4/' + (s.smsPriority ? 'priority' : 'messages');
  var reqs = items.map(function (m) {
    var payload = { apikey: key, number: m.number, message: m.text };
    if (s.smsSender) payload.sendername = s.smsSender;
    return { url: url, method: 'post', payload: payload, muteHttpExceptions: true };
  });
  var resps;
  try { resps = UrlFetchApp.fetchAll(reqs); } catch (e) {
    out.error = 'Could not reach Semaphore: ' + e.message;
    items.forEach(function (m) { out.failed.push(m.seat); });
    return out;
  }
  resps.forEach(function (r, k) {
    var seat = items[k].seat, body = r.getContentText() || '', data = null;
    try { data = JSON.parse(body); } catch (e) { /* not JSON */ }
    var okRow = Array.isArray(data) && data.length && data[0].message_id && String(data[0].status).toLowerCase() !== 'failed';
    if (r.getResponseCode() === 200 && okRow) out.sent.push(seat);
    else {
      out.failed.push(seat);
      if (!out.error) out.error = 'Semaphore (' + r.getResponseCode() + '): ' + body.slice(0, 140);
    }
  });
  return out;
}

function menuSetupSms() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var has = !!props.getProperty('SEMAPHORE_KEY');
  var res = ui.prompt('📱 Set up SMS (Semaphore)',
    '1) Create an account at semaphore.co and load credits.\n2) Copy your API key from the Semaphore dashboard and paste it here.' +
    (has ? '\n\nA key is already saved. Paste a new one to replace it, or type REMOVE to delete it.' : ''), ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var v = res.getResponseText().trim();
  if (!v) return;
  if (v.toUpperCase() === 'REMOVE') { props.deleteProperty('SEMAPHORE_KEY'); ui.alert('SMS key removed. SMS alerts are off.'); return; }
  props.setProperty('SEMAPHORE_KEY', v);
  ui.alert('📱 SMS ready', 'Key saved (kept private — not stored in the sheet).\n\nNext: 📣 Find Me → 📱 Send a test SMS to your own number.\n' +
    'Sender name and other SMS options are in the "' + SET_SHEET + '" tab.', ui.ButtonSet.OK);
}

function menuTestSms() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('📱 Send a test SMS', 'Your mobile number (e.g. 09171234567):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var num = normMobile_(res.getResponseText());
  if (!num) { ui.alert('That doesn\'t look like a PH mobile number (09XXXXXXXXX).'); return; }
  _settings = null;
  var s = getSettings_();
  var out = sendSms_([{ seat: 'A1', number: num, text: 'Find Me sample alert: ' + smsText_(s, 'A1') }], s);
  ui.alert(out.sent.length ? '✅ Sent to ' + num + '. It should arrive within a few seconds.' : '❌ Not sent. ' + (out.error || ''));
}

/** "This is my seat — get staff help": flags a taken seat red so staff can check. */
function claim_(d) {
  var s = getSettings_();
  if (!s.registrationOpen) return { ok: false, code: 'CLOSED' };
  var seat = normSeat_(d.seat);
  var name = String(d.name || '').replace(/\s+/g, ' ').trim().slice(0, 80);
  var device = String(d.deviceId || '');
  if (name.length < 2) return { ok: false, code: 'NAME_REQUIRED' };
  if (!SEAT_RE.test(seat) || !device) return { ok: false, code: 'BAD_REQUEST' };

  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  var result, pending = [];
  try {
    var reg = getRegistry_();
    var i = reg.index[seat];
    if (i === undefined) return { ok: false, code: 'SEAT_NOT_FOUND' };
    var row = reg.rows[i];
    var holder = row[C.NAME - 1];
    // Seat became free (or it's really theirs) → just register.
    if (!holder || row[C.DEVICE - 1] === device || normName_(holder) === normName_(name)) {
      pending = writeRegistration_(reg, i, name, device, d.subscription);
      refreshIssues_(reg);
      invalidate_();
      result = { ok: true, registered: true, seat: seat, status: row[C.STATUS - 1], serverTime: Date.now() };
    } else {
      var existing = parseClaim_(row[C.CLAIM - 1]);
      if (existing && existing.device !== device) return { ok: false, code: 'CLAIM_EXISTS' };
      // Drop this phone's claims elsewhere; free a seat it holds by mistake.
      reg.rows.forEach(function (r, j) {
        if (j === i) return;
        var cl = parseClaim_(r[C.CLAIM - 1]);
        if (cl && cl.device === device) { r[C.CLAIM - 1] = ''; writeRow_(reg, j); }
        if (r[C.DEVICE - 1] === device) pending = pending.concat(releaseSeat_(reg, j, false));
      });
      row[C.CLAIM - 1] = JSON.stringify({
        name: safeText_(name).replace(/^'/, ''), device: device,
        sub: d.subscription ? String(d.subscription).slice(0, 2000) : '', at: Date.now()
      });
      writeRow_(reg, i);
      refreshIssues_(reg);
      paintSeat_(null, row);
      invalidate_();
      result = { ok: true, claimed: true, seat: seat, serverTime: Date.now() };
    }
  } finally {
    lock.releaseLock();
  }
  if (pending.length) sendPush_(pending, getSettings_());
  return result;
}

/** Phone said "I entered the wrong seat" while waiting for staff. */
function unclaim_(d) {
  var seat = normSeat_(d.seat);
  var device = String(d.deviceId || '');
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var reg = getRegistry_();
    var i = reg.index[seat];
    if (i === undefined) return { ok: true };
    var cl = parseClaim_(reg.rows[i][C.CLAIM - 1]);
    if (cl && cl.device === device) {
      reg.rows[i][C.CLAIM - 1] = '';
      writeRow_(reg, i);
      refreshIssues_(reg);
      paintSeat_(null, reg.rows[i]);
      invalidate_();
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function ack_(d) {
  var seat = normSeat_(d.seat);
  var device = String(d.deviceId || '');
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var reg = getRegistry_();
    var i = reg.index[seat];
    if (i === undefined) return { ok: false, code: 'SEAT_NOT_FOUND' };
    var row = reg.rows[i];
    if (row[C.DEVICE - 1] !== device) return { ok: false, code: 'NOT_YOUR_SEAT' };
    if (row[C.STATUS - 1] !== STATUS.ON_THE_WAY) {
      row[C.STATUS - 1] = STATUS.ON_THE_WAY;
      row[C.ACK_AT - 1] = new Date();
      row[C.ALERT - 1] = false;
      writeRow_(reg, i);
      paintSeat_(null, row);
      invalidate_();
    }
    return { ok: true, seat: seat, status: STATUS.ON_THE_WAY };
  } finally {
    lock.releaseLock();
  }
}
