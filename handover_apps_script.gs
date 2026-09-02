/**
 * AHD — Contract → Factory Handover (Google Apps Script, bound to the SALES TRACKER sheet)
 * ---------------------------------------------------------------------------------------
 * WHAT IT DOES
 *   1. The MOMENT a client is ticked CONTRACTED, it emails orders@/crm@/ahdh@ a handover
 *      checklist (no waiting for any daily batch) and stamps the contract date.
 *   2. Every morning it re-reminds on any contracted client whose handover isn't finished,
 *      and flags RED once the order hasn't reached the factory within 3 weeks.
 *   3. The team marks each step done by ticking 4 checkbox columns in THIS sheet:
 *          MS ACCESS · MEP · TECH DWG · RENDERS
 *      A client drops off the reminders automatically once ORDER (sent to factory) is ticked.
 *
 * ONE-TIME SETUP (Ahmed):
 *   Extensions ▸ Apps Script ▸ paste this file ▸ Save ▸ run  setup()  once ▸ authorize.
 *   setup() creates any missing columns, makes the 4 task columns checkboxes, and installs
 *   both triggers (instant-on-edit + daily reminder). Nothing else to configure.
 *
 * Emails are sent from the Google account that authorizes the script (i.e. yours).
 */

// ---------------- CONFIG (edit these if needed) ----------------
var RECIPIENTS   = 'orders@amrhelmydesigns.com, crm@amrhelmydesigns.com, ahdh@amrhelmydesigns.com';
var SLA_DAYS     = 21;                 // order must reach the factory within 3 weeks of contract
var REMINDER_HOUR = 8;                 // daily reminder send hour (sheet's timezone)
var TRACKER_GID  = 57990844;           // the "CLIENT STATUS LIST" tab (pinned so we never
                                       // target another tab that also has a CLIENT header)

// Tracker column headers the script relies on (matched case-insensitively).
var COL = {
  client:    ['CLIENT'],
  rep:       ['SALES PERSON', 'SALESPERSON'],
  amount:    ['Amount in EGP', 'Amount'],
  contracted:['CONTRACTED'],
  order:     ['ORDER'],
  contractDate: ['CONTRACT DATE'],     // auto-created + auto-stamped by this script
};

// The 4 handover tasks → their checkbox column headers (created if missing).
var TASKS = [
  { key: 'MS ACCESS', label: 'Order on the MS Access delivery system' },
  { key: 'MEP',       label: 'MEP drawings started' },
  { key: 'TECH DWG',  label: 'Technical drawings started' },
  { key: 'RENDERS',   label: 'Final renders + presentation sent to the client' },
];

// Room/section words: a row whose CLIENT looks like one of these is a sub-room of the
// client above it, so we walk up to the real client name.
var ROOM_WORDS = ['kitchen','kitchenette','dressing','wardrobe','closet','pantry','laundry',
  'vanity','buffet','reception','living','bedroom','bed room','nanny','maid','tv unit',
  'office','study','bathroom','bath','dining','sofa','cladding','storage','walk in','walk-in',
  'cupboard','cabinet','doors','table','corian','marble','island','entrance','terrace',
  'garden','balcony','roof','مطبخ','دريسنج','غرفة','دولاب','ريسبشن','حمام'];

// ============================================================================
// SETUP — run once
// ============================================================================
function setup() {
  var sh = getTracker_();
  var hr = headerRowIndex_(sh);                 // 1-based header row
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(hr, 1, 1, lastCol).getValues()[0];

  // 1) Ensure every task column + CONTRACT DATE exists; append any that are missing.
  var wanted = TASKS.map(function (t) { return t.key; }).concat(COL.contractDate[0]);
  wanted.forEach(function (name) {
    if (findCol_(headers, [name]) < 0) {
      lastCol += 1;
      sh.getRange(hr, lastCol).setValue(name);
      headers.push(name);
    }
  });

  // 2) Make the 4 task columns checkboxes for the data rows.
  var firstData = hr + 1;
  var nRows = Math.max(sh.getLastRow() - hr, 1);
  TASKS.forEach(function (t) {
    var c = findCol_(headers, [t.key]) + 1;
    sh.getRange(firstData, c, nRows, 1).insertCheckboxes();
  });

  // 3) Install triggers (remove old copies first so re-running setup() is safe).
  removeTriggers_(['onContractEdit', 'dailyHandoverReminders']);
  ScriptApp.newTrigger('onContractEdit')
    .forSpreadsheet(sh.getParent()).onEdit().create();
  ScriptApp.newTrigger('dailyHandoverReminders')
    .timeBased().atHour(REMINDER_HOUR).everyDays(1).create();

  SpreadsheetApp.getActive().toast(
    'Handover automation is live: instant email on CONTRACTED + daily reminders.', 'AHD Handover', 8);
}

// ============================================================================
// INSTANT — fires the moment a client is ticked CONTRACTED
// ============================================================================
function onContractEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getSheetId() !== getTracker_().getSheetId()) return;   // only the tracker tab

    var hr = headerRowIndex_(sh);
    var headers = sh.getRange(hr, 1, 1, sh.getLastColumn()).getValues()[0];
    var cContract = findCol_(headers, COL.contracted);
    if (cContract < 0) return;

    // The edit may span several rows (paste / fill); handle each edited row that
    // touches the CONTRACTED column and was just checked.
    var r0 = e.range.getRow(), c0 = e.range.getColumn();
    var nR = e.range.getNumRows(), nC = e.range.getNumColumns();
    if (c0 > cContract + 1 || c0 + nC - 1 < cContract + 1) return;  // edit didn't include CONTRACTED col

    for (var i = 0; i < nR; i++) {
      var row = r0 + i;
      if (row <= hr) continue;                                     // header/above
      var checked = isChecked_(sh.getRange(row, cContract + 1).getValue());
      if (!checked) continue;
      maybeFireContract_(sh, headers, hr, row);
    }
  } catch (err) {
    // Never let an onEdit error surface to the user editing the sheet.
    console.error('onContractEdit: ' + err);
  }
}

function maybeFireContract_(sh, headers, hr, row) {
  var cDate = findCol_(headers, COL.contractDate);
  if (cDate < 0) return;                                          // setup() not run yet
  var dateCell = sh.getRange(row, cDate + 1);
  if (dateCell.getValue()) return;                               // already stamped ⇒ already emailed

  dateCell.setValue(new Date());                                 // start the 3-week clock now
  var info = rowInfo_(sh, headers, hr, row);
  if (!info.client) return;

  var subject = '🏭 New contract — ' + info.client + ': start factory handover';
  sendMail_(subject, instantHtml_(info));
}

// ============================================================================
// DAILY — re-remind until the handover is done / order sent to factory
// ============================================================================
function dailyHandoverReminders() {
  var sh = getTracker_();
  var hr = headerRowIndex_(sh);
  var lastRow = sh.getLastRow();
  if (lastRow <= hr) return;
  var headers = sh.getRange(hr, 1, 1, sh.getLastColumn()).getValues()[0];
  var cContract = findCol_(headers, COL.contracted);
  var cOrder    = findCol_(headers, COL.order);
  var cDate     = findCol_(headers, COL.contractDate);
  if (cContract < 0 || cOrder < 0) return;

  var vals = sh.getRange(hr + 1, 1, lastRow - hr, sh.getLastColumn()).getValues();
  var pending = [], overdue = 0;

  for (var i = 0; i < vals.length; i++) {
    var rowVals = vals[i];
    if (!isChecked_(rowVals[cContract])) continue;                // not contracted
    if (isChecked_(rowVals[cOrder])) continue;                    // already at the factory ⇒ done

    // Backfill a contract date for deals contracted before this script existed.
    if (cDate >= 0 && !rowVals[cDate]) {
      sh.getRange(hr + 1 + i, cDate + 1).setValue(new Date());
      rowVals[cDate] = new Date();
    }
    var info = rowInfo_(sh, headers, hr, hr + 1 + i, rowVals);
    var remaining = TASKS.filter(function (t) {
      var c = findCol_(headers, [t.key]);
      return c < 0 || !isChecked_(rowVals[c]);
    });
    if (remaining.length === 0) continue;                         // all 4 done ⇒ nothing to nag
    info.remaining = remaining;
    if (info.days > SLA_DAYS) { info.overdue = true; overdue++; }
    pending.push(info);
  }
  if (!pending.length) return;                                    // nothing outstanding today

  // Overdue first, then longest-waiting.
  pending.sort(function (a, b) {
    return (a.overdue === b.overdue) ? (b.days - a.days) : (a.overdue ? -1 : 1);
  });
  var subject = '🏭 Contract handover — ' + pending.length + ' pending'
              + (overdue ? (' · ' + overdue + ' OVERDUE 🔴') : '');
  sendMail_(subject, reminderHtml_(pending));
}

// ============================================================================
// EMAIL HTML
// ============================================================================
function instantHtml_(info) {
  var tasks = TASKS.map(function (t) {
    return '<tr><td style="padding:6px 10px;font-size:14px">⬜ ' + esc_(t.label) + '</td></tr>';
  }).join('');
  return wrap_(
    '<h2 style="margin:0 0 4px;color:#1c1c1c">New contract signed — start the factory handover</h2>'
    + '<p style="margin:0 0 12px;color:#555;font-size:13px"><b>' + esc_(info.client) + '</b>'
    + (info.rep ? ' · ' + esc_(info.rep) : '')
    + (info.amount ? ' · ' + esc_(info.amount) + ' EGP' : '') + '</p>'
    + '<p style="font-size:13px;color:#333;margin:0 0 8px">Please action now:</p>'
    + '<table style="border-collapse:collapse;background:#fbf7ee;border:1px solid #ece4d2;border-radius:8px;width:100%">'
    + tasks + '</table>'
    + '<p style="font-size:12px;color:#8a6a12;margin:12px 0 0"><b>Also remind the sales team</b> to send the final renders + presentation to the client.</p>'
    + slaLine_(0)
    + tickHint_());
}

function reminderHtml_(pending) {
  var rows = pending.map(function (info) {
    var checks = TASKS.map(function (t) {
      var isDone = !info.remaining.some(function (x) { return x.key === t.key; });
      return '<div style="font-size:12px;line-height:1.6;color:' + (isDone ? '#1f7a44' : '#b06a12') + '">'
        + (isDone ? '✅ ' : '⬜ ') + esc_(t.label) + '</div>';
    }).join('');
    var sla = info.overdue
      ? '<span style="color:#c0392b;font-weight:700;font-size:12px">🔴 ' + info.days + 'd since contract — OVERDUE (&gt;' + SLA_DAYS + 'd to factory)</span>'
      : '<span style="color:#8a8a8a;font-size:12px">' + info.days + 'd since contract · ' + (SLA_DAYS - info.days) + 'd left to factory</span>';
    return '<tr style="border-bottom:1px solid #f2ede2' + (info.overdue ? ';background:#fdecec' : '') + '">'
      + '<td style="padding:8px 10px;vertical-align:top">'
      + '<b>' + esc_(info.client) + '</b>'
      + '<div style="font-size:11px;color:#9a9a9a">' + (info.rep ? esc_(info.rep) + ' · ' : '')
      + (info.amount ? esc_(info.amount) + ' EGP' : '') + '</div>'
      + '<div style="margin-top:3px">' + sla + '</div></td>'
      + '<td style="padding:8px 10px;vertical-align:top">' + checks + '</td></tr>';
  }).join('');
  return wrap_(
    '<h2 style="margin:0 0 8px;color:#1c1c1c">Contracts awaiting factory handover</h2>'
    + '<table style="border-collapse:collapse;width:100%;border:1px solid #ece4d2;border-radius:8px">'
    + rows + '</table>'
    + tickHint_());
}

function slaLine_(days) {
  return '<p style="font-size:12px;color:#8a8a8a;margin:10px 0 0">Target: the order should reach the factory (tick <b>ORDER</b>) within <b>' + SLA_DAYS + ' days</b> of the contract.</p>';
}

function tickHint_() {
  var url = getTracker_().getParent().getUrl();
  return '<p style="font-size:12px;color:#8a8a8a;margin:10px 0 0">Mark each step done by ticking '
    + '<b>MS ACCESS · MEP · TECH DWG · RENDERS</b> in the '
    + '<a href="' + url + '" style="color:#8a6a12;font-weight:600">sales tracker</a>. '
    + 'This clears automatically once <b>ORDER</b> (sent to factory) is ticked.</p>';
}

function wrap_(inner) {
  return '<div style="font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;max-width:640px;'
    + 'margin:0 auto;padding:16px;background:#fff;color:#2b2b2b">' + inner
    + '<p style="font-size:10px;color:#a9a190;margin-top:16px">Auto-sent by the AHD tracker handover automation.</p></div>';
}

// ============================================================================
// HELPERS
// ============================================================================
function getTracker_() {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  // Prefer the pinned tab by its gid so we can never target another tab that also
  // happens to carry a CLIENT header.
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === TRACKER_GID) return sheets[i];
  }
  // Fallback: the first tab that carries a CLIENT header (in case the gid changed).
  for (var j = 0; j < sheets.length; j++) {
    try { if (headerRowIndex_(sheets[j], true) > 0) return sheets[j]; } catch (e) {}
  }
  return ss.getActiveSheet();
}

// 1-based row that carries the CLIENT header (searches the first 5 rows). If probe
// is true, returns -1 instead of defaulting, so getTracker_ can pick the right tab.
function headerRowIndex_(sh, probe) {
  var n = Math.min(5, sh.getLastRow());
  if (n < 1) return probe ? -1 : 1;
  var vals = sh.getRange(1, 1, n, sh.getLastColumn()).getValues();
  for (var r = 0; r < n; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      if (String(vals[r][c]).trim().toUpperCase() === 'CLIENT') return r + 1;
    }
  }
  return probe ? -1 : 1;
}

// 0-based column index for the first header matching any name (case-insensitive), or -1.
function findCol_(headers, names) {
  var up = names.map(function (n) { return String(n).trim().toUpperCase(); });
  for (var i = 0; i < headers.length; i++) {
    if (up.indexOf(String(headers[i]).trim().toUpperCase()) >= 0) return i;
  }
  return -1;
}

function isChecked_(v) {
  return v === true || String(v).trim().toUpperCase() === 'TRUE';
}

function isSubroom_(name) {
  var n = String(name || '').trim().toLowerCase();
  if (!n) return true;
  for (var i = 0; i < ROOM_WORDS.length; i++) if (n.indexOf(ROOM_WORDS[i]) >= 0) return true;
  return false;
}

// Build a client info object for a row (walking up to the parent if it's a sub-room).
function rowInfo_(sh, headers, hr, row, rowVals) {
  var cClient = findCol_(headers, COL.client);
  var cRep    = findCol_(headers, COL.rep);
  var cAmt    = findCol_(headers, COL.amount);
  var cDate   = findCol_(headers, COL.contractDate);
  if (!rowVals) rowVals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];

  var client = cClient >= 0 ? String(rowVals[cClient]).trim() : '';
  var r = row;
  // Walk up while the client cell is blank or a sub-room name.
  while (r > hr + 1 && (!client || isSubroom_(client))) {
    r -= 1;
    var up = sh.getRange(r, 1, 1, sh.getLastColumn()).getValues()[0];
    var c = cClient >= 0 ? String(up[cClient]).trim() : '';
    if (c && !isSubroom_(c)) { client = c; rowVals = up; break; }
  }
  var amt = cAmt >= 0 ? Number(String(rowVals[cAmt]).replace(/[^0-9.]/g, '')) : 0;
  var days = 0;
  if (cDate >= 0 && rowVals[cDate]) {
    var d = new Date(rowVals[cDate]);
    days = Math.floor((new Date() - d) / 86400000);
  }
  return {
    client: client,
    rep: cRep >= 0 ? String(rowVals[cRep]).trim() : '',
    amount: amt ? Math.round(amt).toLocaleString('en-US') : '',
    days: days, overdue: false, remaining: [],
  };
}

function sendMail_(subject, htmlBody) {
  MailApp.sendEmail({ to: RECIPIENTS, subject: subject, htmlBody: htmlBody, noReply: false });
}

function removeTriggers_(names) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (names.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
}

function esc_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
