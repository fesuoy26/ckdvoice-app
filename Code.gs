/**
 * Chokkad Voice Pravasi Koottayma — read-only data service for the web app.
 * Reg. No. MPM/218/2026
 *
 * Endpoints (deploy as Web app → Execute as: Me → Access: Anyone):
 *   GET  ?action=summary                      → aggregate figures only (no names)
 *   GET  ?action=check&no=&name=&mobile=      → one member result, or "not found"
 *   POST {"action":"check","no":"","name":"","mobile":""}   (same as above; used by the app
 *        so that names and phone numbers never appear in a URL)
 *
 * Columns are found BY HEADER NAME, so adding columns or rows never breaks the app.
 * Nothing in this file ever returns the member list.
 */

// ───────────── SETTINGS (edit these if a sheet is renamed) ─────────────
var CONFIG = {
  MEMBER_SHEET: 'MEMBERSHIP2027',      // membership register
  KH_SHEET: 'KARUNYA HASTHAM',         // Karunya Hastham loan register ("KH.Register")
  KH_SUMMARY_SHEET: 'KH SUMMARY',      // optional: "Fund allocation" is read from here
  TIMEZONE: 'Asia/Riyadh',             // validity is checked against today's date here
  RENEWAL_YEAR: 2027,                  // used only if the named range "RenewalYear" is missing
  KH_FUND_FALLBACK: 300000,            // used only if "Fund allocation" is not found
  KUDUMBA_PER_MEMBER: 500,             // ₹ per valid member (Kudumba Suraksha)
  KUDUMBA_TARGET: 100000,              // ₹ family support target
  CACHE_SECONDS: 300,                  // 5 minutes
  CONTACT: '+91 9400237218'
};

// Header names in the membership sheet (case and extra spaces are ignored).
var MEMBER_HEADERS = {
  no: 'MEMB. NO.',
  name: 'NAME',
  house: 'HOUSE NAME',
  mobile: 'CONT. NUMBER',
  upTo: 'UP TO',              // current validity date
  dueDate: 'DATE',            // first "DATE" column = next payment due date
  dueAmount: 'AMOUNT',        // amount to pay for the renewal period
  newValidity: 'VALIDITY'     // validity after paying the amount above
};
// Optional status column. If none of these headers exists, a text note typed
// into the next-due "DATE" column (as the register does today) marks a special account.
var STATUS_HEADERS = ['STATUS', 'REMARKS', 'STATUS/REMARKS'];

// ───────────── WEB ENTRY POINTS ─────────────
function doGet(e) {
  var p = (e && e.parameter) || {};
  return route_(p);
}

function doPost(e) {
  var p = {};
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { p = {}; }
  return route_(p);
}

function route_(p) {
  var out;
  try {
    var action = String(p.action || 'summary').toLowerCase();
    if (action === 'summary') out = getSummary_();
    else if (action === 'check') out = checkMember_(p.no, p.name, p.mobile);
    else out = { ok: false, error: 'unknown_action' };
  } catch (err) {
    out = { ok: false, error: 'server_error' };   // never echo internal details
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// ───────────── SUMMARY ─────────────
function getSummary_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('summary_v1');
  if (hit) return JSON.parse(hit);

  var reg = loadMembers_();
  var today = todayStr_();
  var renewalEnd = reg.renewalYear + '-12-31';
  var valid = 0, expired = 0, expiring = 0;
  reg.members.forEach(function (m) {
    if (m.special) return;                       // deceased / suspended / etc. are not counted
    if (m.upTo && m.upTo >= today) {
      valid++;
      if (m.upTo < renewalEnd) expiring++;
    } else {
      expired++;
    }
  });
  var dots = [];                                 // sorted: valid first, so no member can be identified
  for (var i = 0; i < valid; i++) dots.push(1);
  for (var j = 0; j < expired; j++) dots.push(0);

  var out = {
    ok: true,
    updated: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd-MM-yyyy HH:mm'),
    asOf: today,
    renewalYear: reg.renewalYear,
    members: { valid: valid, expired: expired, total: valid + expired, expiringSoon: expiring },
    dots: dots,
    kudumba: {
      perMember: CONFIG.KUDUMBA_PER_MEMBER,
      target: CONFIG.KUDUMBA_TARGET,
      pool: valid * CONFIG.KUDUMBA_PER_MEMBER
    },
    kh: getKH_()
  };
  cache.put('summary_v1', JSON.stringify(out), CONFIG.CACHE_SECONDS);
  return out;
}

// Karunya Hastham — aggregate figures only. No names, numbers or member IDs leave this function.
function getKH_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.KH_SHEET);
  if (!sh) return null;
  var v = sh.getDataRange().getValues();
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

  var hr = findHeaderRow_(v, ['MEMB. #', 'AMOUNT PAID']);
  if (hr < 0) return null;
  var h = v[hr].map(norm_);
  var cMemb = h.indexOf('MEMB. #');
  var cLoan = h.indexOf('AMOUNT PAID');
  var cLoanDate = cLoan + 1;
  var cBal = h.indexOf('BALANCE');
  var cStatus = h.indexOf('STATUS');
  var repay = [];                                   // every "AMOUNT" column after the loan, with its DATE
  for (var c = cLoan + 1; c < h.length; c++) if (h[c] === 'AMOUNT') repay.push(c);

  var loans = 0, open = 0, closed = 0, disbursed = 0, repaid = 0;
  var people = {}, years = {};
  function yr(y) { if (!years[y]) years[y] = { year: +y, disbursed: 0, repaid: 0 }; return years[y]; }

  for (var r = hr + 1; r < v.length; r++) {
    var row = v[r];
    var memb = Number(row[cMemb]);
    var loan = Math.abs(Number(row[cLoan]) || 0);
    if (!(memb >= 1000) || !loan) continue;         // skips totals, notes and blank rows
    loans++;
    people[memb] = 1;
    disbursed += loan;
    var ly = yearOf_(row[cLoanDate], tz);
    if (ly) yr(ly).disbursed += loan;
    var paidThis = 0;
    repay.forEach(function (rc) {
      var a = Number(row[rc]) || 0;
      if (!a) return;
      paidThis += a;
      var y = yearOf_(row[rc + 1], tz);
      if (y) yr(y).repaid += a;
    });
    repaid += paidThis;
    var st = cStatus >= 0 ? norm_(row[cStatus]) : '';
    var bal = cBal >= 0 ? Math.abs(Number(row[cBal]) || 0) : loan - paidThis;
    if (st === 'OPEN' || (!st && bal > 0)) open++; else closed++;
  }

  var fund = readLabelValue_(CONFIG.KH_SUMMARY_SHEET, 'FUND ALLOCATION') || CONFIG.KH_FUND_FALLBACK;
  var outstanding = disbursed - repaid;
  var list = Object.keys(years).sort().map(function (k) { return years[k]; });
  return {
    loans: loans,
    beneficiaries: Object.keys(people).length,
    open: open,
    closed: closed,
    disbursed: disbursed,
    repaid: repaid,
    outstanding: outstanding,
    repaymentPct: disbursed ? Math.round(repaid / disbursed * 1000) / 10 : 0,
    fund: fund,
    available: fund - outstanding,
    years: list
  };
}

// ───────────── MEMBERSHIP CHECK ─────────────
function checkMember_(no, name, mobile) {
  var qNo = String(no || '').replace(/\D/g, '');
  var qName = normName_(name);
  var qMob = String(mobile || '').replace(/\D/g, '');
  qMob = qMob.length >= 9 ? qMob.slice(-9) : '';
  var given = (qNo ? 1 : 0) + (qName.length >= 3 ? 1 : 0) + (qMob ? 1 : 0);
  if (given < 2) return { ok: true, found: false, reason: 'need_two', contact: CONFIG.CONTACT };

  var reg = loadMembers_();
  var best = [], bestScore = 0;
  reg.members.forEach(function (m) {
    var s = 0;
    if (qNo && String(m.no) === qNo) s++;
    if (qName.length >= 3 && nameMatches_(qName, m)) s++;
    if (qMob && m.mob9 && m.mob9 === qMob) s++;
    if (s >= 2) {
      if (s > bestScore) { bestScore = s; best = [m]; }
      else if (s === bestScore) best.push(m);
    }
  });
  if (best.length !== 1) return { ok: true, found: false, contact: CONFIG.CONTACT };

  var m = best[0];
  if (m.special) {
    return {
      ok: true, found: true, status: 'special', no: m.no,
      message: 'ഈ അംഗത്വ അക്കൗണ്ട് നിലവിൽ സജീവമല്ല. കൂടുതൽ വിവരങ്ങൾക്ക് ഭാരവാഹികളുമായി ബന്ധപ്പെടുക.',
      contact: CONFIG.CONTACT
    };
  }
  var today = todayStr_();
  var renewalEnd = reg.renewalYear + '-12-31';
  var status = !m.upTo || m.upTo < today ? 'expired' : (m.upTo < renewalEnd ? 'expiring' : 'valid');
  var res = {
    ok: true, found: true, status: status,
    no: m.no, name: m.name, house: m.house,
    validity: dmy_(m.upTo),
    contact: CONFIG.CONTACT
  };
  if (status === 'valid') {
    res.next = { date: dmy_(m.dueDate), amount: m.dueAmount, until: dmy_(m.newValidity) };
  } else {
    res.due = { amount: m.dueAmount, from: dmy_(m.dueDate), until: dmy_(m.newValidity) };
  }
  return res;
}

function nameMatches_(q, m) {
  var full = normName_(m.name);
  var withHouse = normName_(m.name + ' ' + m.house);
  var houseFirst = normName_(m.house + ' ' + m.name);
  if (q === full || q === withHouse || q === houseFirst) return true;
  // first name typed alone, or name typed with only part of the house name
  return full.length >= 3 && (q.indexOf(full) === 0 || full.indexOf(q) === 0);
}

// ───────────── REGISTER READER (cached, server-side only) ─────────────
function loadMembers_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('members_v1');
  if (hit) return JSON.parse(hit);

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CONFIG.MEMBER_SHEET);
  if (!sh) throw new Error('Member sheet not found');
  var tz = ss.getSpreadsheetTimeZone();
  var v = sh.getDataRange().getValues();
  var hr = findHeaderRow_(v, [MEMBER_HEADERS.no, MEMBER_HEADERS.name]);
  if (hr < 0) throw new Error('Header row not found');
  var h = v[hr].map(norm_);
  var col = {};
  Object.keys(MEMBER_HEADERS).forEach(function (k) { col[k] = h.indexOf(norm_(MEMBER_HEADERS[k])); });
  var cStatus = -1;
  STATUS_HEADERS.forEach(function (s) { if (cStatus < 0) cStatus = h.indexOf(norm_(s)); });

  var members = [];
  for (var r = hr + 1; r < v.length; r++) {
    var row = v[r];
    var no = Number(row[col.no]);
    if (!(no >= 1000)) continue;                 // ignores explanatory note rows (< 1000) and blanks
    var dueCell = col.dueDate >= 0 ? row[col.dueDate] : '';
    var statusText = cStatus >= 0 ? String(row[cStatus] || '').trim() : '';
    var special = !!statusText || (typeof dueCell === 'string' && dueCell.trim() !== '' && !parseDate_(dueCell, tz));
    var mob = String(row[col.mobile] || '').replace(/\D/g, '');
    members.push({
      no: no,
      name: String(row[col.name] || '').trim(),
      house: col.house >= 0 ? String(row[col.house] || '').trim() : '',
      mob9: mob.length >= 9 ? mob.slice(-9) : '',
      upTo: parseDate_(row[col.upTo], tz),
      dueDate: special ? null : parseDate_(dueCell, tz),
      dueAmount: special ? null : (Number(row[col.dueAmount]) || 0),
      newValidity: special ? null : parseDate_(row[col.newValidity], tz),
      special: special
    });
  }
  var out = { members: members, renewalYear: renewalYear_(ss) };
  try { cache.put('members_v1', JSON.stringify(out), CONFIG.CACHE_SECONDS); } catch (e) { /* too large: skip cache */ }
  return out;
}

function renewalYear_(ss) {
  try {
    var rg = ss.getRangeByName('RenewalYear');
    var y = rg && Number(rg.getValue());
    if (y > 2000 && y < 2100) return y;
  } catch (e) { }
  return CONFIG.RENEWAL_YEAR;
}

// ───────────── HELPERS ─────────────
function norm_(s) { return String(s == null ? '' : s).toUpperCase().replace(/\s+/g, ' ').trim(); }
function normName_(s) { return String(s || '').toUpperCase().replace(/[^A-Z]/g, ''); }
function todayStr_() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd'); }

function findHeaderRow_(v, needed) {
  var want = needed.map(norm_);
  for (var r = 0; r < Math.min(v.length, 15); r++) {
    var h = v[r].map(norm_);
    if (want.every(function (w) { return h.indexOf(w) >= 0; })) return r;
  }
  return -1;
}

// Returns 'yyyy-MM-dd' or null. Accepts real dates and "dd/mm/yyyy" text.
function parseDate_(x, tz) {
  if (x instanceof Date && !isNaN(x)) return Utilities.formatDate(x, tz, 'yyyy-MM-dd');
  var m = String(x || '').trim().match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return null;
}
function yearOf_(x, tz) { var d = parseDate_(x, tz); return d ? d.slice(0, 4) : null; }
function dmy_(iso) { return iso ? iso.slice(8, 10) + '-' + iso.slice(5, 7) + '-' + iso.slice(0, 4) : null; }

function readLabelValue_(sheetName, label) {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return 0;
  var v = sh.getDataRange().getValues();
  var want = norm_(label);
  for (var r = 0; r < v.length; r++) {
    for (var c = 0; c < v[r].length - 1; c++) {
      if (norm_(v[r][c]) === want) {
        for (var k = c + 1; k < v[r].length; k++) { var n = Number(v[r][k]); if (n) return n; }
      }
    }
  }
  return 0;
}

// Run this once from the editor (▶ Run) to clear the 5-minute cache after big edits.
function clearCache() {
  CacheService.getScriptCache().removeAll(['summary_v1', 'members_v1']);
}
