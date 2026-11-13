/**
 * Code.js — GST bill automation, LOCAL-ONLY edition (Web App).
 *
 * FLOW (no tunnel, no daily trigger, no server running on the PC):
 *   You run  python run_billing.py  on the PC whenever you want to bill.
 *   That script calls THIS web app twice:
 *
 *     1. action "get_due"
 *        Reads UPCOMING BILLS. A row is DUE when Expire Date <= a cutoff
 *        (sent by Python; default = yesterday). For each due row it:
 *          - finds the customer's LAST invoice in GST RECORDS (fuzzy match)
 *            so its GSTIN/description/amounts/payment mode can be copied,
 *          - computes the next GLOBAL invoice no from the last row
 *            (ST/046/2026-27 -> ST/047/2026-27; restarts ST/001/2027-28
 *            after each 31 March),
 *          - returns the list to Python.
 *        This app NEVER creates or writes the "Bill File" / "Bill Status"
 *        columns (Bill File is only ever READ, if the column exists).
 *
 *     2. action "complete"
 *        Python edited the .docx files and made the PDFs, then sends back
 *        per-row results. For each success this app appends the new row to
 *        GST RECORDS (Payment Status "Pending", date as text) — ONLY if
 *        that invoice number is not already there (crash-safe) — then
 *        ADVANCES the customer's Expire Date to the new period end,
 *        UPDATES the Duration column (e.g. "1 Year" / "3 Months", sent by
 *        Python from the new bill period) and RE-SORTS UPCOMING BILLS by
 *        Expire Date (nearest first).
 *
 * SECRETS: SECRET / SHEET_ID come from Script Properties (Project Settings),
 *          seeded once via setupSecrets() or rotated via the "set_props"
 *          action — never hardcode them in this file.
 *
 * ONE-TIME SETUP:
 *   1. Seed secrets (see setupSecrets below) with the .env values.
 *   2. Deploy > New deployment > type: Web app
 *        Execute as: Me       Who has access: Anyone
 *      Copy the Web App URL (ends in /exec) into .env:
 *        appscript_url :- https://script.google.com/macros/s/.../exec
 *   3. In the editor run dryRunOnce() (authorize when asked) — preview only.
 *   4. On the PC:  python run_billing.py --dry-run   then without --dry-run.
 *
 * NOTE: after editing this code you must create a NEW DEPLOYMENT VERSION
 *       (Deploy > Manage deployments > Edit > Version: New), or simply run
 *       `python deploy_apps_script.py` on the PC, or the old code keeps
 *       running.
 */

var PROPS = PropertiesService.getScriptProperties();
// Secrets live in Script Properties (Apps Script: Project Settings), seeded /
// rotated from .env values via POST {action:'set_props', secret, props}.
// First-time setup: run setupSecrets() in the editor once (values pasted there).
var SHEET_ID = PROPS.getProperty('SHEET_ID') || '';
var SECRET = PROPS.getProperty('SECRET') || '';
var UPCOMING_SHEET = 'UPCOMING BILLS';
var GST_SHEET = 'GST RECORDS';

var COL_BILL_FILE = 'Bill File';     // UPCOMING BILLS: optional exact docx name (READ ONLY)
var COL_DURATION = 'Duration';       // UPCOMING BILLS: written after billing, e.g. '1 Year' / '3 Months'
var PREFIX = 'ST';                   // invoice prefix

var INVOICE_RE = /^ST\/(\d+)\/(\d{4})-(\d{1,2})$/;
var GST_COLUMNS = ['Invoice No', 'Invoice Date', 'Customer Name', 'Customer GSTIN',
  'Service Description', 'Amount ₹', 'CGST 9% (₹)', 'SGST 9% (₹)', 'Total Amount (₹)',
  'Payment Mode', 'Payment Status', 'Payment Date', 'GST PAID'];
var STOPWORDS = ['PVT', 'LTD', 'LLP', 'PRIVATE', 'LIMITED', 'CO', 'THE', 'AND',
  'BILL', 'SOLUTIONS', 'ENTERPRISES', 'INDIA', 'STORES'];

/** Web app entry point. Python POSTs JSON {action, secret, ...}. */
function doPost(e) {
  var body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return jsonOut_({ ok: false, error: 'bad JSON body' });
  }
  if (!SECRET)
    return jsonOut_({ ok: false,
      error: 'server not configured — seed Script Properties first (see setupSecrets in the editor or POST set_props)' });
  if (String(body.secret || '') !== SECRET)
    return jsonOut_({ ok: false, error: 'bad or missing secret' });
  try {
    if (body.action === 'get_due') return jsonOut_(getDue_(body));
    if (body.action === 'complete') return jsonOut_(complete_(body));
    if (body.action === 'set_props') return jsonOut_(setProps_(body));
    return jsonOut_({ ok: false, error: 'unknown action: ' + body.action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  }
}

/** Browser test that the web app is alive. */
function doGet() {
  return jsonOut_({ ok: true, service: 'gst-bill-automation',
                    actions: ['get_due', 'complete', 'set_props'] });
}

/**
 * One-time setup, run manually in the Apps Script editor (▶ button):
 * replace the two placeholder arguments with the real values from the
 * local .env (appscript_secret / sheet_id), run it, authorize, done.
 * Never commit real values here — after seeding, restore the placeholders.
 */
function setupSecrets() {
  var secret = '';   // <-- paste .env appscript_secret
  var sheetId = '';  // <-- paste .env sheet_id
  var p = PropertiesService.getScriptProperties();
  if (secret) p.setProperty('SECRET', secret);
  if (sheetId) p.setProperty('SHEET_ID', sheetId);
  Logger.log('configured: SECRET=' + (p.getProperty('SECRET') ? 'yes' : 'no')
    + '  SHEET_ID=' + (p.getProperty('SHEET_ID') ? 'yes' : 'no'));
}

/**
 * Secret-gated property management (rotation / first-time seeding).
 * POST {action:'set_props', secret:<current>, props:{SECRET:'…', SHEET_ID:'…'}}
 * Never echoes values back; returns only the property names that were set.
 */
function setProps_(body) {
  var allowed = { SECRET: 1, SHEET_ID: 1 };
  var set = [];
  var props = body.props || {};
  Object.keys(props).forEach(function (k) {
    if (!allowed[k]) return;
    var v = String(props[k] || '').trim();
    if (!v) return;
    PROPS.setProperty(k, v);
    set.push(k);
  });
  if (!set.length) throw new Error('no valid properties to set (allowed: SECRET, SHEET_ID)');
  if (set.indexOf('SECRET') >= 0) SECRET = PROPS.getProperty('SECRET');
  if (set.indexOf('SHEET_ID') >= 0) SHEET_ID = PROPS.getProperty('SHEET_ID');
  return { ok: true, set: set };
}

/** Editor-only preview: logs exactly what get_due would return (marks nothing). */
function dryRunOnce() {
  var r = getDue_({ dry: true, secret: SECRET });
  Logger.log('cutoff: ' + r.cutoff + '   due: ' + r.due.length);
  r.due.forEach(function (d) {
    Logger.log('  ' + d.customer + ' | expires ' + d.expire +
      ' | next ' + d.new_invoice_no + ' | last ' + (d.last_invoice_no || 'NONE') +
      ' | bill_file ' + (d.bill_file || '(auto)'));
  });
  return r;
}

/* -------------------------------------------------------------- action 1 */

function getDue_(body) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var up = ss.getSheetByName(UPCOMING_SHEET);
  var gst = ss.getSheetByName(GST_SHEET);
  if (!up || !gst) throw new Error('tabs "' + UPCOMING_SHEET + '" / "' + GST_SHEET + '" not found');

  var today = startOfDay_(new Date());
  var cutoff = asDate_(body.cutoff) || addDays_(today, -1);
  var upVals = up.getDataRange().getValues();
  var upHeader = upVals[0];
  var iName = colOf_(upHeader, 'Customer name');
  var iExp = colOf_(upHeader, 'Expire Date');
  if (iName < 0 || iExp < 0) throw new Error('UPCOMING BILLS needs "Customer name" and "Expire Date" columns');
  var iFile = colOf_(upHeader, COL_BILL_FILE); // READ ONLY; may not exist — never created

  var gstVals = gst.getDataRange().getValues();
  var due = [];
  for (var r = 1; r < upVals.length; r++) {
    var name = String(upVals[r][iName] || '').trim();
    var expire = asDate_(upVals[r][iExp]);
    if (!name || !expire) continue;
    if (expire > cutoff) continue;

    var last = findLastGstRecord_(gstVals, name);
    due.push({
      key: 'up_r' + (r + 1) + '_' + fmt_(expire),
      row: r + 1,
      customer: name,
      expire: fmt_(expire),
      bill_file: iFile >= 0 ? String(upVals[r][iFile] || '').trim() : '',
      new_invoice_no: nextInvoiceNo_(gstVals, today),
      last_invoice_no: last ? last.invoiceNo : null,
      has_gst_record: !!last
    });
  }

  return { ok: true, dry: !!body.dry, cutoff: fmt_(cutoff), due: due };
}

/* -------------------------------------------------------------- action 2 */

function complete_(body) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var up = ss.getSheetByName(UPCOMING_SHEET);
  var gst = ss.getSheetByName(GST_SHEET);
  var upVals = up.getDataRange().getValues();
  var iName = colOf_(upVals[0], 'Customer name');
  var iExp = colOf_(upVals[0], 'Expire Date');
  if (iName < 0 || iExp < 0)
    throw new Error('UPCOMING BILLS needs "Customer name" and "Expire Date" columns');
  var iDur = colOf_(upVals[0], COL_DURATION);  // written only if the column exists
  var gCols = {};
  var gstVals = gst.getDataRange().getValues();
  GST_COLUMNS.forEach(function (c) { gCols[c] = colOf_(gstVals[0], c); });

  var existing = {};                       // invoice numbers already in GST RECORDS
  for (var r = 1; r < gstVals.length; r++) {
    var v = String(gstVals[r][0] || '').trim();
    if (INVOICE_RE.test(v)) existing[v] = true;
  }

  var log = [];
  var moved = 0;
  var results = body.results || [];
  for (var i = 0; i < results.length; i++) {
    var res = results[i];
    var m = String(res.key || '').match(/^up_r(\d+)_/);
    if (!m) { log.push('bad key ' + res.key); continue; }
    var row = parseInt(m[1], 10) - 1;      // 0-based data index in UPCOMING BILLS
    if (row < 1 || row >= upVals.length) { log.push('row out of range for ' + res.key); continue; }
    var name = String(upVals[row][iName] || '').trim();

    if (!res.ok) {
      // nothing written to the row: it keeps its old Expire Date, so the
      // next run picks it up again automatically (built-in retry).
      log.push(name + ': ERROR ' + String(res.error || '').slice(0, 80));
      continue;
    }

    var invNo = String(res.new_invoice_no || '').trim();
    if (!existing[invNo]) {
      var last = findLastGstRecord_(gstVals, name);
      if (!last) {
        log.push(name + ': no GST record to copy — NOT appended, fill manually');
        continue;
      }
      appendGstRecord_(gst, gCols, {
        invoiceNo: invNo,
        invoiceDate: String(res.invoice_date || '').replace(/-/g, '/'),
        customer: last.customer, gstin: last.gstin, desc: last.desc,
        amount: last.amount, cgst: last.cgst, sgst: last.sgst, total: last.total,
        mode: last.mode
      });
      existing[invNo] = true;
    }

    // advance the customer's Expire Date to the new period end, and record
    // the plan Duration (months sent by Python from the new bill period)
    var note = ' — ' + invNo;
    var newExp = asDate_(res.new_expire);
    if (newExp) {
      var oldExp = asDate_(upVals[row][iExp]);
      var cell = up.getRange(row + 1, iExp + 1);
      cell.setNumberFormat('dd/mm/yy');          // real date, same style as the column
      cell.setValue(newExp);
      upVals[row][iExp] = fmt_(newExp);
      moved++;

      var dm = parseInt(res.duration_months, 10);
      if (!(dm > 0) && oldExp) dm = monthsBetween_(oldExp, newExp); // fallback: expire shift
      if (iDur >= 0 && dm > 0) {
        var durText = durationText_(dm);
        up.getRange(row + 1, iDur + 1).setValue(durText);
        note += ', ' + durText;
      }
      note += ', expire moved to ' + fmt_(newExp);
    } else {
      note += ' recorded (no new expire sent)';
    }
    log.push(name + ': ' + (existing[invNo] && res.duplicate ? 'already recorded' : 'record appended') +
      note + (res.pdf_error ? ' (PDF FAILED)' : ''));
  }

  if (moved) sortUpcomingByExpire_(up, iExp);
  return { ok: true, log: log };
}

/** Re-orders the data rows of UPCOMING BILLS by Expire Date, nearest first. */
function sortUpcomingByExpire_(up, iExp) {
  var lastRow = up.getLastRow();
  if (lastRow < 2) return;
  var lastCol = up.getLastColumn();
  var vals = up.getRange(2, 1, lastRow - 1, lastCol).getValues();
  vals.sort(function (a, b) {
    var da = asDate_(a[iExp]);
    var db = asDate_(b[iExp]);
    if (!da && !db) return 0;
    if (!da) return 1;                         // rows without a date go last
    if (!db) return -1;
    return da - db;
  });
  up.getRange(2, 1, lastRow - 1, lastCol).setValues(vals);
  SpreadsheetApp.flush();
}

/* ------------------------------------------------------ duration utilities */

/** Whole calendar months between two dates (02/01/2026 -> 01/01/2027 = 12). */
function monthsBetween_(d1, d2) {
  return (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth());
}

/** 12 -> '1 Year', 24 -> '2 Years', 1 -> '1 Month', 3 -> '3 Months'. */
function durationText_(m) {
  if (m === 12) return '1 Year';
  if (m > 12 && m % 12 === 0) return (m / 12) + ' Years';
  return m === 1 ? '1 Month' : m + ' Months';
}

/* ------------------------------------------------------- invoice numbering */

var lastInvoiceCache_ = null;

function nextInvoiceNo_(gstVals, billDate) {
  var fyNow = fyOf_(billDate);
  if (lastInvoiceCache_) {                                  // sequential within one request
    var n = (fyNow === lastInvoiceCache_.fy) ? lastInvoiceCache_.num + 1 : 1;
    lastInvoiceCache_ = { num: n, fy: fyNow, width: lastInvoiceCache_.width || 3 };
    return PREFIX + '/' + pad_(n, lastInvoiceCache_.width) + '/' + fyNow;
  }
  // first call of this execution: take the numerically HIGHEST invoice no for
  // the current FY (safe even if the sheet's rows are not in invoice order)
  var maxNum = 0, width = 3;
  for (var r = 1; r < gstVals.length; r++) {
    var m = String(gstVals[r][0] || '').trim().match(INVOICE_RE);
    if (!m || (m[2] + '-' + m[3]) !== fyNow) continue;
    var num = parseInt(m[1], 10);
    if (num > maxNum) { maxNum = num; width = m[1].length; }
  }
  var next = maxNum + 1;
  lastInvoiceCache_ = { num: next, fy: fyNow, width: width };
  return PREFIX + '/' + pad_(next, width) + '/' + fyNow;
}

function fyOf_(d) {
  var y = d.getFullYear();
  if (d.getMonth() + 1 < 4) y -= 1;                         // FY starts 1 April
  return y + '-' + pad_((y + 1) % 100, 2);
}

/* -------------------------------------------------------- GST record store */

function findLastGstRecord_(gstVals, name) {
  var best = null, bestScore = 0, bestDate = null;
  for (var r = 1; r < gstVals.length; r++) {
    var row = gstVals[r];
    var inv = String(row[0] || '').trim();
    if (!INVOICE_RE.test(inv)) continue;
    var s = nameScore_(name, String(row[2] || ''));
    if (s <= 0) continue;
    var d = asDate_(row[1]);
    if (!best || s > bestScore || (s === bestScore && d && bestDate && d > bestDate)) {
      best = { invoiceNo: inv, customer: String(row[2]).trim(), gstin: String(row[3] || '').trim(),
               desc: String(row[4] || '').trim(), amount: row[5], cgst: row[6], sgst: row[7],
               total: row[8], mode: String(row[9] || '').trim() };
      bestScore = s; bestDate = d;
    }
  }
  return best;
}

function appendGstRecord_(gst, gCols, v) {
  var rowIdx = gst.getLastRow() + 1;
  var row = new Array(GST_COLUMNS.length).fill('');
  row[gCols['Invoice No']] = v.invoiceNo;
  row[gCols['Invoice Date']] = v.invoiceDate;
  row[gCols['Customer Name']] = v.customer;
  row[gCols['Customer GSTIN']] = v.gstin;
  row[gCols['Service Description']] = v.desc;
  row[gCols['Amount ₹']] = v.amount;
  row[gCols['CGST 9% (₹)']] = v.cgst;
  row[gCols['SGST 9% (₹)']] = v.sgst;
  row[gCols['Total Amount (₹)']] = v.total;
  row[gCols['Payment Mode']] = v.mode;
  row[gCols['Payment Status']] = 'Pending';
  var range = gst.getRange(rowIdx, 1, 1, GST_COLUMNS.length);
  range.setValues([row]);
  var dCell = gst.getRange(rowIdx, gCols['Invoice Date'] + 1);
  dCell.setNumberFormat('@');                               // keep dd/mm/yy as text
  dCell.setValue(v.invoiceDate);
  SpreadsheetApp.flush();
}

function nameScore_(a, b) {
  var na = normName_(a), nb = normName_(b);
  if (!na || !nb) return 0;
  if (na === nb) return 1;
  if (nb.indexOf(na) === 0 || na.indexOf(nb) === 0) return 0.9;
  var ta = tokens_(a), tb = tokens_(b);
  var sub = ta.length > 0 && ta.every(function (t) { return tb.indexOf(t) >= 0; });
  var sub2 = tb.length > 0 && tb.every(function (t) { return ta.indexOf(t) >= 0; });
  if (sub || sub2) return 0.8;
  var dsc = dice_(na, nb);                                  // fuzzy (typos etc.)
  return dsc >= 0.55 ? dsc : 0;                             // below this = unrelated
}

function normName_(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function tokens_(s) {
  return String(s || '').toUpperCase().split(/[^A-Z0-9]+/)
    .filter(function (t) { return t && STOPWORDS.indexOf(t) < 0; });
}

function dice_(a, b) {
  if (a.length < 2 || b.length < 2) return (a === b) ? 1 : 0;
  var map = {}, hits = 0;
  for (var i = 0; i < a.length - 1; i++) { var g = a.substr(i, 2); map[g] = (map[g] || 0) + 1; }
  for (var j = 0; j < b.length - 1; j++) { var h = b.substr(j, 2); if (map[h] > 0) { hits++; map[h]--; } }
  return (2 * hits) / (a.length - 1 + b.length - 1);
}

/* ------------------------------------------------------------------ helpers */

function colOf_(header, title) {
  for (var i = 0; i < header.length; i++) {
    if (String(header[i]).trim().toLowerCase() === title.toLowerCase()) return i;
  }
  return -1;
}

function asDate_(v) {
  if (v instanceof Date) return startOfDay_(v);
  var s = String(v || '').trim();
  var m = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
  if (!m) return null;
  var y = parseInt(m[3], 10); if (y < 100) y += 2000;
  return new Date(y, parseInt(m[2], 10) - 1, parseInt(m[1], 10));
}

function startOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addDays_(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
function fmt_(d) { return pad_(d.getDate(), 2) + '/' + pad_(d.getMonth() + 1, 2) + '/' + pad_(d.getFullYear() % 100, 2); }
function pad_(n, w) { var s = String(n); while (s.length < w) s = '0' + s; return s; }

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
