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
var COL_ADDRESS = 'Address';         // UPCOMING BILLS: optional buyer address used by the mobile PDF generator
var COL_USERID = 'User id';          // UPCOMING BILLS: optional broadband user id shown on the bill
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
    // ---- mobile app (PWA) actions ----
    if (body.action === 'ping') return jsonOut_(ping_(body));
    if (body.action === 'list_all') return jsonOut_(listAll_(body));
    if (body.action === 'bill_customer') return jsonOut_(billCustomer_(body));
    if (body.action === 'make_pdf') return jsonOut_(makePdf_(body));
    if (body.action === 'debug_pdf') return jsonOut_(debugPdf_(body));
    return jsonOut_({ ok: false, error: 'unknown action: ' + body.action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  }
}

/** Browser test that the web app is alive. */
function doGet() {
  return jsonOut_({ ok: true, service: 'gst-bill-automation',
                    actions: ['get_due', 'complete', 'set_props',
                              'ping', 'list_all', 'bill_customer', 'make_pdf'] });
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
  var allowed = { SECRET: 1, SHEET_ID: 1, BILL_TEMPLATE_DOC_ID: 1 };
  var set = [];
  var props = body.props || {};
  Object.keys(props).forEach(function (k) {
    if (!allowed[k]) return;
    var v = String(props[k] || '').trim();
    if (!v) return;
    PROPS.setProperty(k, v);
    set.push(k);
  });
  if (!set.length) throw new Error('no valid properties to set (allowed: SECRET, SHEET_ID, BILL_TEMPLATE_DOC_ID)');
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

/* =====================================================================
 *  MOBILE APP (PWA) ACTIONS
 *
 *  The phone app POSTs the same {action, secret, ...} JSON (with
 *  Content-Type text/plain, which avoids a CORS preflight).
 *
 *    ping           -> {ok, service}                       liveness check
 *    list_all       -> every UPCOMING BILLS customer (expire date, duration,
 *                      bill file, address, user id) with their FULL invoice
 *                      history from GST RECORDS, sorted nearest expiry first.
 *    bill_customer  -> the whole billing flow in ONE call, in a SAFE ORDER:
 *                      resolve customer -> next global invoice no ->
 *                      GENERATE THE BILL PDF FIRST (in memory — if this
 *                      fails, NOTHING is written to the sheet) -> only after
 *                      the PDF succeeded: append GST RECORDS row (Payment
 *                      Status Pending, crash-safe dedupe) -> advance Expire
 *                      Date -> update Duration/Address/User id columns ->
 *                      re-sort UPCOMING BILLS -> return everything incl.
 *                      the PDF as base64 (the PHONE keeps it in local
 *                      storage; Google Drive is never touched).
 *    make_pdf       -> regenerate the PDF for an invoice that already
 *                      exists in GST RECORDS (e.g. a bill whose PDF was
 *                      never delivered, or was deleted on the phone).
 *                      Sheet data only — no Drive, no sheet writes.
 * ===================================================================== */

function ping_(body) {
  return { ok: true, service: 'gst-bill-automation', time: new Date().toISOString() };
}

/* ------------------------------------------------------------------ list_all */

function listAll_(body) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var up = ss.getSheetByName(UPCOMING_SHEET);
  var gst = ss.getSheetByName(GST_SHEET);
  if (!up || !gst) throw new Error('tabs "' + UPCOMING_SHEET + '" / "' + GST_SHEET + '" not found');

  var upVals = up.getDataRange().getValues();
  var gstVals = gst.getDataRange().getValues();
  var iName = colOf_(upVals[0], 'Customer name');
  var iExp = colOf_(upVals[0], 'Expire Date');
  if (iName < 0 || iExp < 0) throw new Error('UPCOMING BILLS needs "Customer name" and "Expire Date" columns');
  var iDur = colOf_(upVals[0], COL_DURATION);
  var iFile = colOf_(upVals[0], COL_BILL_FILE);
  var iAddr = colOf_(upVals[0], COL_ADDRESS);
  var iUser = colOf_(upVals[0], COL_USERID);

  var gCols = {};
  GST_COLUMNS.forEach(function (c) { gCols[c] = colOf_(gstVals[0], c); });

  // collect ALL gst rows once, grouped later per customer by fuzzy name
  var gstRows = [];
  for (var g = 1; g < gstVals.length; g++) {
    var inv = String(gstVals[g][0] || '').trim();
    if (!INVOICE_RE.test(inv)) continue;
    gstRows.push({
      invoice_no: inv,
      invoice_date: gCols['Invoice Date'] >= 0 ? String(gstVals[g][gCols['Invoice Date']] || '').trim() : '',
      customer: gCols['Customer Name'] >= 0 ? String(gstVals[g][gCols['Customer Name']] || '').trim() : '',
      gstin: gCols['Customer GSTIN'] >= 0 ? String(gstVals[g][gCols['Customer GSTIN']] || '').trim() : '',
      desc: gCols['Service Description'] >= 0 ? String(gstVals[g][gCols['Service Description']] || '').trim() : '',
      amount: num_(gstVals[g][gCols['Amount ₹']]),
      cgst: num_(gstVals[g][gCols['CGST 9% (₹)']]),
      sgst: num_(gstVals[g][gCols['SGST 9% (₹)']]),
      total: num_(gstVals[g][gCols['Total Amount (₹)']]),
      mode: gCols['Payment Mode'] >= 0 ? String(gstVals[g][gCols['Payment Mode']] || '').trim() : '',
      status: gCols['Payment Status'] >= 0 ? String(gstVals[g][gCols['Payment Status']] || '').trim() : '',
      payment_date: gCols['Payment Date'] >= 0 ? String(gstVals[g][gCols['Payment Date']] || '').trim() : '',
      gst_paid: gCols['GST PAID'] >= 0 ? String(gstVals[g][gCols['GST PAID']] || '').trim() : '',
      _date: asDate_(gstVals[g][gCols['Invoice Date']])
    });
  }
  gstRows.sort(function (a, b) { return (b._date ? b._date.getTime() : 0) - (a._date ? a._date.getTime() : 0); });

  var customers = [];
  for (var r = 1; r < upVals.length; r++) {
    var name = String(upVals[r][iName] || '').trim();
    var expire = asDate_(upVals[r][iExp]);
    if (!name) continue;

    // fuzzy-match this customer against every gst row (same scoring the
    // desktop billing uses to copy data), keep score > 0, newest first
    var invoices = [];
    for (var k = 0; k < gstRows.length; k++) {
      if (nameScore_(name, gstRows[k].customer) > 0) invoices.push(gstRows[k]);
    }
    var lastInv = invoices.length ? invoices[0] : null;

    customers.push({
      row: r + 1,
      name: name,
      expire: expire ? fmtDMY_(expire) : '',
      expire_iso: expire ? iso_(expire) : '',
      duration: iDur >= 0 ? String(upVals[r][iDur] || '').trim() : '',
      bill_file: iFile >= 0 ? String(upVals[r][iFile] || '').trim() : '',
      address: iAddr >= 0 ? String(upVals[r][iAddr] || '').trim() : '',
      userid: iUser >= 0 ? String(upVals[r][iUser] || '').trim() : '',
      gstin: lastInv ? lastInv.gstin : '',
      last_invoice: lastInv ? stripRow_(lastInv) : null,
      invoices: invoices.map(stripRow_)
    });
  }
  customers.sort(function (a, b) {
    if (!a.expire_iso && !b.expire_iso) return 0;
    if (!a.expire_iso) return 1;
    if (!b.expire_iso) return -1;
    return a.expire_iso < b.expire_iso ? -1 : (a.expire_iso > b.expire_iso ? 1 : 0);
  });

  return { ok: true, today: fmtDMY_(startOfDay_(new Date())), today_iso: iso_(startOfDay_(new Date())), customers: customers };
}

function stripRow_(v) {
  return { invoice_no: v.invoice_no, invoice_date: v.invoice_date, customer: v.customer,
           gstin: v.gstin, desc: v.desc, amount: v.amount, cgst: v.cgst, sgst: v.sgst,
           total: v.total, mode: v.mode, status: v.status, payment_date: v.payment_date,
           gst_paid: v.gst_paid };
}

function num_(v) {
  var n = parseFloat(String(v === null || v === undefined ? '' : v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function iso_(d) {
  return d.getFullYear() + '-' + pad_(d.getMonth() + 1, 2) + '-' + pad_(d.getDate(), 2);
}

/* ------------------------------------------------------------- bill_customer */

function billCustomer_(body) {
  var name = String(body.customer || '').trim();
  if (!name) throw new Error('customer is required');

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var up = ss.getSheetByName(UPCOMING_SHEET);
  var gst = ss.getSheetByName(GST_SHEET);
  if (!up || !gst) throw new Error('tabs "' + UPCOMING_SHEET + '" / "' + GST_SHEET + '" not found');

  var upVals = up.getDataRange().getValues();
  var iName = colOf_(upVals[0], 'Customer name');
  var iExp = colOf_(upVals[0], 'Expire Date');
  if (iName < 0 || iExp < 0) throw new Error('UPCOMING BILLS needs "Customer name" and "Expire Date" columns');
  var iDur = colOf_(upVals[0], COL_DURATION);
  var iAddr = colOf_(upVals[0], COL_ADDRESS);
  var iUser = colOf_(upVals[0], COL_USERID);

  // --- resolve the row: exact name first, then unique fuzzy match ----------
  var row = -1;
  for (var r = 1; r < upVals.length; r++) {
    if (String(upVals[r][iName] || '').trim().toLowerCase() === name.toLowerCase()) { row = r; break; }
  }
  var candidates = [];
  if (row < 0) {
    for (var r2 = 1; r2 < upVals.length; r2++) {
      var s = nameScore_(name, String(upVals[r2][iName] || ''));
      if (s >= 0.8) candidates.push({ row: r2, score: s, name: String(upVals[r2][iName] || '').trim() });
    }
    if (candidates.length === 1) row = candidates[0].row;
  }
  if (row < 0) {
    var list = candidates.slice(0, 8).map(function (c) { return c.name; }).join(' | ');
    throw new Error('customer not found uniquely in UPCOMING BILLS: ' + name + (list ? ' — matches: ' + list : ''));
  }
  var custName = String(upVals[row][iName] || '').trim();

  var oldExp = asDate_(upVals[row][iExp]);
  if (!oldExp) throw new Error('customer has no valid Expire Date: ' + custName);

  // --- last gst record (source of gstin / description / payment mode) ------
  var gstVals = gst.getDataRange().getValues();
  var last = findLastGstRecord_(gstVals, custName);
  if (!last) throw new Error('no GST RECORDS entry to copy for ' + custName + ' — add one invoice row manually first');

  // --- duration -----------------------------------------------------------
  var months = parseInt(body.duration_months, 10);
  if (!(months > 0)) months = parseDuration_(iDur >= 0 ? String(upVals[row][iDur] || '') : '');
  if (!(months > 0)) throw new Error('duration_months not sent and Duration column of ' + custName + ' is not parseable');

  // --- period + invoice no -------------------------------------------------
  var today = startOfDay_(new Date());
  var periodStart = addDays_(oldExp, 1);
  var periodEnd = addDays_(addMonths_(periodStart, months), -1);
  var invoiceNo = nextInvoiceNo_(gstVals, today);

  // --- amounts -------------------------------------------------------------
  var amount = parseFloat(body.amount);
  if (!(amount > 0)) amount = last.amount > 0 ? last.amount : 0;
  if (!(amount > 0)) throw new Error('no amount known for ' + custName + ' — send amount in the request');
  var cgst = round2_(amount * 0.09);
  var sgst = round2_(amount * 0.09);
  var total = round2_(amount + cgst + sgst);
  var desc = String(body.description || '').trim() || last.desc;
  var mode = String(body.payment_mode || '').trim() || last.mode || 'Cash';

  // --- buyer extras (read now, written to the sheet only after PDF OK) -----
  var address = String(body.address || '').trim();
  if (!address && iAddr >= 0) address = String(upVals[row][iAddr] || '').trim();
  var userid = String(body.userid || '').trim();
  if (!userid && iUser >= 0) userid = String(upVals[row][iUser] || '').trim();

  // --- generate the PDF FIRST (in memory only — NEVER saved to Drive) ------
  // Nothing is written to the sheet unless this succeeds.
  var bill = {
    buyer: last.customer,
    address: address,
    gstin: last.gstin,
    userid: userid,
    invoiceNo: invoiceNo,
    invoiceDate: fmtShort_(today),
    periodStart: fmtDMY_(periodStart),
    periodEnd: fmtDMY_(periodEnd),
    desc: desc,
    amount: amount, cgst: cgst, sgst: sgst, total: total,
    mode: mode
  };
  var pdf;
  try {
    pdf = renderBillPdf_(bill);
  } catch (err) {
    throw new Error('PDF generation failed — sheet NOT updated. ' + String(err && err.message || err));
  }
  var pdfBlob = pdf.blob;

  // --- PDF OK -> only now update the sheet ---------------------------------
  // (append GST row + advance expiry + duration/address/user id + re-sort)
  var sheetError = null;
  var appended = false;
  try {
    var existing = {};
    for (var g = 1; g < gstVals.length; g++) {
      var v = String(gstVals[g][0] || '').trim();
      if (INVOICE_RE.test(v)) existing[v] = true;
    }
    if (!existing[invoiceNo]) {
      var gCols = {};
      GST_COLUMNS.forEach(function (c) { gCols[c] = colOf_(gstVals[0], c); });
      appendGstRecord_(gst, gCols, {
        invoiceNo: invoiceNo,
        invoiceDate: fmtShort_(today),
        customer: last.customer, gstin: last.gstin, desc: desc,
        amount: amount, cgst: cgst, sgst: sgst, total: total, mode: mode
      });
      appended = true;
    }

    var cell = up.getRange(row + 1, iExp + 1);
    cell.setNumberFormat('dd/mm/yy');
    cell.setValue(periodEnd);
    if (iDur >= 0) up.getRange(row + 1, iDur + 1).setValue(durationText_(months));

    if (address && iAddr < 0) iAddr = ensureTextColumn_(up, COL_ADDRESS);
    if (address && iAddr >= 0) up.getRange(row + 1, iAddr + 1).setValue(address);

    if (userid && iUser < 0) iUser = ensureTextColumn_(up, COL_USERID);
    if (userid && iUser >= 0) up.getRange(row + 1, iUser + 1).setValue(userid);

    sortUpcomingByExpire_(up, iExp);
  } catch (err) {
    sheetError = String(err && err.message || err);
  }

  return {
    ok: true,
    customer: custName,
    invoice_no: invoiceNo,
    invoice_date: fmtShort_(today),
    appended: appended,
    sheet_updated: sheetError === null,
    sheet_error: sheetError,
    duration_months: months,
    period: { start: fmtDMY_(periodStart), end: fmtDMY_(periodEnd) },
    new_expire: fmtDMY_(periodEnd),
    amount: amount, cgst: cgst, sgst: sgst, total: total,
    desc: desc, mode: mode,
    pdf_source: pdf.source,
    pdf_name: custName.replace(/[\\\/:*?"<>|]+/g, '_').replace(/\s+/g, '_') + '_BILL.pdf',
    pdf_base64: Utilities.base64Encode(pdfBlob.getBytes())
  };
}

/* -------------------------------------------------------------------- make_pdf */

/**
 * Regenerates the bill PDF for an invoice that already exists in GST
 * RECORDS. No sheet writes, no Drive — the PDF is built in memory from the
 * stored row and returned as base64 (the phone saves it locally).
 * Period start/end are reconstructed from the customer's CURRENT expire
 * date (= that bill's period end) minus the stored Duration.
 */
function makePdf_(body) {
  var invNo = String(body.invoice_no || '').trim();
  if (!invNo) throw new Error('invoice_no is required');

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var gst = ss.getSheetByName(GST_SHEET);
  var up = ss.getSheetByName(UPCOMING_SHEET);
  if (!gst || !up) throw new Error('tabs "' + UPCOMING_SHEET + '" / "' + GST_SHEET + '" not found');

  var vals = gst.getDataRange().getValues();
  var gCols = {};
  GST_COLUMNS.forEach(function (c) { gCols[c] = colOf_(vals[0], c); });

  var gRow = -1;
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][0] || '').trim() === invNo) { gRow = r; break; }
  }
  if (gRow < 0) throw new Error('invoice not found in GST RECORDS: ' + invNo);

  var customer = gCols['Customer Name'] >= 0 ? String(vals[gRow][gCols['Customer Name']] || '').trim() : '';
  if (!customer) throw new Error('invoice row has no Customer Name: ' + invNo);

  // buyer extras (address / user id) from UPCOMING BILLS by fuzzy name
  var upVals = up.getDataRange().getValues();
  var iName = colOf_(upVals[0], 'Customer name');
  var iAddr = colOf_(upVals[0], COL_ADDRESS);
  var iUser = colOf_(upVals[0], COL_USERID);
  var iDur = colOf_(upVals[0], COL_DURATION);
  var iExp = colOf_(upVals[0], 'Expire Date');
  var address = '', userid = '', durationText = '', expire = null, best = 0;
  for (var u = 1; u < upVals.length; u++) {
    var s = nameScore_(customer, String(upVals[u][iName] || ''));
    if (s > best) {
      best = s;
      address = iAddr >= 0 ? String(upVals[u][iAddr] || '').trim() : '';
      userid = iUser >= 0 ? String(upVals[u][iUser] || '').trim() : '';
      durationText = iDur >= 0 ? String(upVals[u][iDur] || '').trim() : '';
      expire = iExp >= 0 ? asDate_(upVals[u][iExp]) : null;
    }
  }

  // reconstruct the bill period: current expire = that bill's period end
  var months = parseDuration_(durationText);
  var periodEnd = expire || null;
  var periodStart = (periodEnd && months > 0) ? addDays_(addMonths_(periodEnd, -months), 1) : null;

  var bill = {
    buyer: customer,
    address: address,
    gstin: gCols['Customer GSTIN'] >= 0 ? String(vals[gRow][gCols['Customer GSTIN']] || '').trim() : '',
    userid: userid,
    invoiceNo: invNo,
    invoiceDate: gCols['Invoice Date'] >= 0 ? String(vals[gRow][gCols['Invoice Date']] || '').trim() : '',
    periodStart: periodStart ? fmtDMY_(periodStart) : '-',
    periodEnd: periodEnd ? fmtDMY_(periodEnd) : '-',
    desc: gCols['Service Description'] >= 0 ? String(vals[gRow][gCols['Service Description']] || '').trim() : '',
    amount: num_(vals[gRow][gCols['Amount ₹']]),
    cgst: num_(vals[gRow][gCols['CGST 9% (₹)']]),
    sgst: num_(vals[gRow][gCols['SGST 9% (₹)']]),
    total: num_(vals[gRow][gCols['Total Amount (₹)']]),
    mode: gCols['Payment Mode'] >= 0 ? String(vals[gRow][gCols['Payment Mode']] || '').trim() : ''
  };
  var pdf = renderBillPdf_(bill);

  return {
    ok: true,
    invoice_no: invNo,
    customer: customer,
    pdf_source: pdf.source,
    pdf_name: customer.replace(/[\\\/:*?"<>|]+/g, '_').replace(/\s+/g, '_') + '_BILL.pdf',
    pdf_base64: Utilities.base64Encode(pdf.blob.getBytes())
  };
}

/* ------------------------------------------------------------- pdf template */

/** Replicates the Shreeganesh Telecom .docx bill layout as HTML (A4). */
function buildBillHtml_(v) {
  var rupee = '&#8377;';
  var rows = [
    ['Invoice Date', v.invoiceDate],
    ['Invoice No', v.invoiceNo],
    ['User id', v.userid || '-'],
    ['Support', '9373322868 / 9822333443'],
    ['PAN', 'APNPA7087H'],
    ['GSTIN / UIN', '27APNPA7087H1ZK'],
    ['Place of Supply', 'MAHARASHTRA'],
    ['State Code', '27'],
    ['E-Mail', 'shreeganeshtelecom14@gmail.com']
  ].map(function (p) {
    return '<tr><td class="lbl">' + p[0] + ' :-</td><td class="val">' + escapeHtml_(p[1]) + '</td></tr>';
  }).join('');

  var addrLines = String(v.address || '').split('\n')
    .map(function (l) { return escapeHtml_(l.trim()); }).filter(Boolean)
    .map(function (l) { return '<div>' + l + '</div>'; }).join('');

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    '@page { size: A4; margin: 0; }' +
    'body { font-family: Arial, Helvetica, sans-serif; font-size: 12px; color:#000; width: 760px; margin: 0 auto; padding: 24px 18px; }' +
    'table { border-collapse: collapse; width: 100%; }' +
    'td, th { vertical-align: top; }' +
    '.hdr { margin-bottom: 6px; }' +
    '.co { font-size: 21px; font-weight: bold; letter-spacing: .5px; }' +
    '.co2 { font-size: 15px; font-weight: bold; }' +
    '.coaddr { font-size: 12px; margin-top: 3px; }' +
    '.taxinv { text-align: center; font-size: 19px; font-weight: bold; padding-left: 40px; }' +
    '.main td { padding: 6px 4px; }' +
    '.buyer-h { font-weight: bold; font-size: 12px; }' +
    '.buyer { font-size: 12px; line-height: 1.45; }' +
    '.meta .lbl { white-space: nowrap; padding-right: 10px; font-size: 12px; }' +
    '.meta .val { font-size: 12px; font-weight: bold; }' +
    '.period { font-size: 12.5px; font-weight: bold; margin: 10px 0 8px; }' +
    '.items, .items td { border: 1px solid #000; font-size: 12px; }' +
    '.items th { border: 1px solid #000; font-size: 12px; padding: 5px; text-align: center; }' +
    '.items td { padding: 6px 8px; }' +
    '.items .c { text-align: center; } .items .r { text-align: right; }' +
    '.totals { margin-top: 6px; }' +
    '.totals td { padding: 4px 8px; font-size: 12px; }' +
    '.totals .r { text-align: right; width: 110px; }' +
    '.totrow td { font-weight: bold; border-top: 1px solid #000; border-bottom: 1px solid #000; font-size: 13px; }' +
    '.words { font-size: 12px; font-weight: bold; margin-top: 10px; }' +
    '.foot { margin-top: 14px; }' +
    '.foot td { font-size: 11px; vertical-align: top; padding: 0 6px; }' +
    '.sig { margin-top: 8px; font-weight: bold; font-size: 11px; }' +
    '.signspace { height: 42px; }' +
    '.bank b { display: inline-block; min-width: 130px; font-weight: normal; }' +
    '</style></head><body>' +

    '<table class="hdr"><tr>' +
    '<td style="width:55%"><span class="co">SHREEGANESH TELECOM</span><br>' +
    '<span class="co2">PUNE SPEEDNET BROADBAND</span><br>' +
    '<span class="coaddr">Address:- 396, Gawaliwada, Khadki, Pune - 411003.</span></td>' +
    '<td class="taxinv">TAX INVOICE</td></tr></table>' +

    '<table class="main"><tr>' +
    '<td style="width:46%">' +
    '<div class="buyer-h">Buyer (Bill to)</div><div class="buyer">' +
    '<div style="margin-top:8px">' + escapeHtml_(v.buyer) + '</div>' +
    addrLines +
    '<div>GSTIN / UIN : - ' + escapeHtml_(v.gstin || '-') + '</div>' +
    '<div>State Name : - Maharashtra, Code: 27</div>' +
    '<div>SAC Code :- 9984</div>' +
    '</div></td>' +
    '<td><table class="meta">' + rows + '</table></td>' +
    '</tr></table>' +

    '<div class="period">Bill Period:-' + v.periodStart + ' to ' + v.periodEnd + '</div>' +

    '<table class="items"><tr>' +
    '<th style="width:7%">SR<br>No</th><th style="width:53%">Plan<br>OTC (One Time Charges)</th>' +
    '<th style="width:20%">Charges (' + rupee + ')</th><th style="width:20%">Amount (' + rupee + ')</th></tr>' +
    '<tr><td class="c">1.</td><td>' + escapeHtml_(v.desc || '-') + '</td>' +
    '<td class="r">' + money_(v.amount) + '</td><td class="r">' + money_(v.amount) + '</td></tr></table>' +

    '<table class="totals" style="width:62%; margin-left:38%">' +
    '<tr><td>Basic Amount</td><td class="r">' + money_(v.amount) + '</td></tr>' +
    '<tr><td>CGST - 9%</td><td class="r">' + money_(v.cgst) + '</td></tr>' +
    '<tr><td>SGST - 9%</td><td class="r">' + money_(v.sgst) + '</td></tr>' +
    '<tr><td>Round off</td><td class="r">' + money_(0) + '</td></tr>' +
    '<tr class="totrow"><td>Total (' + rupee + ')</td><td class="r">' + money_(v.total) + '</td></tr>' +
    '</table>' +

    '<div class="words">Amount Chargeable (in words): ' + rupeeWords_(v.total) + ' RUPEES ONLY/_</div>' +

    '<table class="foot"><tr>' +
    '<td style="width:52%"><b>Terms &amp; Conditions</b><br><br>' +
    '1) Interest (18%) will be chargeable if invoice(s) past due 15 days. ' +
    'Overdue, Non-Payment may lead to suspend service. Client has to make the ' +
    'payment within the 15 Days Reserved time.<br>' +
    '2) All the bank charges, dishonor cheque charges if any are born by customer a/c</td>' +
    '<td><b>Company\'s Bank Details</b><br><div class="bank">' +
    'A/c Holder\'s Name  :  SHREEGANESH TELECOM<br>' +
    'Bank Name&nbsp;&nbsp;: Bank Of India - 050420110000602<br>' +
    'A/c No.&nbsp;&nbsp;: 050420110000602<br>' +
    'Branch &amp; IFS Code&nbsp;&nbsp;: BKID0000504</div>' +
    '<div class="signspace"></div><div class="sig" style="text-align:right">FOR SHREEGANESH TELECOM</div></td>' +
    '</tr></table>' +

    '</body></html>';
}

function escapeHtml_(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function money_(n) {
  return (Math.round((parseFloat(n) || 0) * 100) / 100).toFixed(2);
}

/* ------------------------------------------------- Word-template rendering */

/** Diagnostic action: tries every step of the Word-template PDF pipeline
 *  separately and reports which step fails. Read-only; the copy it makes is
 *  trashed immediately. */
function debugPdf_(body) {
  var steps = [];
  function step(name, fn) {
    try {
      var v = fn();
      steps.push({ step: name, ok: true, info: v === undefined ? '' : String(v) });
    } catch (e) {
      steps.push({ step: name, ok: false, error: String(e && e.message || e).slice(0, 160) });
    }
  }
  var docId = billTemplateDocId_();
  step('template docId', function () { return docId || 'NOT SET'; });
  step('sheet access', function () { return SpreadsheetApp.openById(SHEET_ID).getName(); });
  step('DriveApp.getFileById', function () {
    if (!docId) throw new Error('no docId set');
    return 'file: ' + DriveApp.getFileById(docId).getName();
  });
  var copyId = '';
  step('makeCopy', function () {
    if (!docId) throw new Error('no docId set');
    var c = DriveApp.getFileById(docId).makeCopy('debug-copy');
    copyId = c.getId();
    return 'copy: ' + c.getName();
  });
  step('DocumentApp.openById(copy)', function () {
    if (!copyId) throw new Error('no copy made');
    return DocumentApp.openById(copyId).getBody().getParagraphs().length + ' paragraphs';
  });
  step('copy getAs(pdf)', function () {
    if (!copyId) throw new Error('no copy made');
    return DocumentApp.openById(copyId).getAs('application/pdf').getBytes().length + ' bytes';
  });
  if (copyId) step('trash copy', function () { DriveApp.getFileById(copyId).setTrashed(true); });
  step('doc content', function () {
    if (!docId) throw new Error('no docId set');
    var d = DocumentApp.openById(docId);
    return d.getImages().length + ' images (header/footers), ' +
           d.getBody().getTables().length + ' tables, ' +
           d.getBody().getParagraphs().length + ' paragraphs';
  });
  step('ORIGINAL doc getAs(pdf)', function () {
    if (!docId) throw new Error('no docId set');
    return DocumentApp.openById(docId).getAs('application/pdf').getBytes().length + ' bytes';
  });
  step('html getAs(pdf)', function () {
    return Utilities.newBlob('<p>x</p>', 'text/html', 'x.html').getAs('application/pdf').getBytes().length + ' bytes';
  });
  return { ok: true, debug: steps };
}

/** Editor test: run this function (not doGet!) to trigger the consent dialog
 *  for the documents + drive scopes, and verify the bill template doc is
 *  reachable. Read-only, changes nothing. */
function testTemplateAccess() {
  var id = billTemplateDocId_();
  if (!id) { Logger.log('no template doc id registered'); return; }
  var name = DriveApp.getFileById(id).getName();
  var body = DocumentApp.openById(id).getBody();
  Logger.log('template OK: "' + name + '"  body paragraphs: ' + body.getParagraphs().length);
}

/** Google Doc id of the bill template (set once via set_props / test_backend).
 *  It must be the GOOGLE DOC created from GST_BILL_TEMPLATE.docx ("Open with
 *  Google Docs" converts the uploaded docx and preserves the layout). */
function billTemplateDocId_() {
  return PROPS.getProperty('BILL_TEMPLATE_DOC_ID') || '';
}

/** Renders the bill PDF. Prefers the Word template (identical to the .docx
 *  bills); falls back to the built-in HTML layout when no template is set
 *  or the template fails. */
function renderBillPdf_(bill) {
  var docId = billTemplateDocId_();
  if (docId) {
    try {
      return { blob: renderBillTemplatePdf_(docId, bill), source: 'word-template' };
    } catch (e) {
      try {
        return { blob: htmlBillPdf_(bill),
                 source: 'html-fallback (' + String(e && e.message || e).slice(0, 200) + ')' };
      } catch (e2) {
        throw new Error('PDF render failed — template: ' +
          String(e && e.message || e).slice(0, 80) + ' | html: ' +
          String(e2 && e2.message || e2).slice(0, 80));
      }
    }
  }
  return { blob: htmlBillPdf_(bill), source: 'html' };
}

function htmlBillPdf_(bill) {
  return Utilities.newBlob(buildBillHtml_(bill), 'text/html', 'bill.html').getAs('application/pdf');
}

/** Renders one bill from the Word template. A THROWAWAY COPY of the template
 *  doc is made, filled with the {{tokens}}, exported as PDF and deleted in
 *  the same execution — no bill content is ever kept in Drive, and
 *  concurrent bills can never collide. */
function renderBillTemplatePdf_(docId, bill) {
  var vals = templateValues_(bill);
  var copy = DriveApp.getFileById(docId).makeCopy('bill-' + bill.invoiceNo);
  try {
    var doc = DocumentApp.openById(copy.getId());
    var targets = [doc.getBody(), doc.getHeader(), doc.getFooter()];
    for (var t = 0; t < targets.length; t++) {
      var el = targets[t];
      if (!el) continue;
      for (var tok in vals) el.replaceText(tok, vals[tok]);
    }
    return doc.getAs('application/pdf');
  } finally {
    try { copy.setTrashed(true); } catch (e) { /* keep going */ }
  }
}

/** Token -> value map matching GST_BILL_TEMPLATE.docx (make_template.py). */
function templateValues_(bill) {
  var desc = String(bill.desc || '');
  var line1 = desc, line2 = '';
  var nl = desc.indexOf('\n');
  if (nl >= 0) {
    line1 = desc.slice(0, nl).trim();
    line2 = desc.slice(nl + 1).trim();
  } else if (desc.length > 46) {
    // two visual lines in the template — split a long description at a space
    var mid = desc.lastIndexOf(' ', Math.round(desc.length / 2));
    if (mid > 20) { line1 = desc.slice(0, mid); line2 = desc.slice(mid + 1); }
  }
  // tokens inside values would corrupt later replacements — neutralise them
  function clean(v) {
    return String(v === null || v === undefined ? '' : v)
      .replace(/\{\{/g, '{ {').replace(/\}\}/g, '} }');
  }
  var addr = String(bill.address || '').split('\n')
    .map(function (l) { return l.trim(); }).filter(Boolean).join(', ');
  var vals = {};
  vals['{{BUYER}}'] = clean(bill.buyer);
  vals['{{ADDRESS}}'] = clean(addr);
  vals['{{GSTIN}}'] = clean(bill.gstin);
  vals['{{USERID}}'] = clean(bill.userid || '-');
  vals['{{INVOICE_NO}}'] = clean(bill.invoiceNo);
  vals['{{INVOICE_DATE}}'] = clean(String(bill.invoiceDate || '').replace(/\//g, '-'));
  vals['{{PERIOD}}'] = clean(bill.periodStart + ' to ' + bill.periodEnd);
  vals['{{DESC1}}'] = clean(line1);
  vals['{{DESC2}}'] = clean(line2);
  vals['{{AMOUNT}}'] = money_(bill.amount);
  vals['{{CGST}}'] = money_(bill.cgst);
  vals['{{ROUND}}'] = '00.00';
  vals['{{TOTAL}}'] = money_(bill.total);
  vals['{{WORDS}}'] = clean(rupeeWords_(bill.total) + ' RUPEES ONLY/_');
  return vals;
}

/* ------------------------------------------------------------- pdf utilities */

/** Ensures UPCOMING BILLS has a plain-text column; returns its 0-based index. */
function ensureTextColumn_(up, title) {
  var header = up.getDataRange().getValues()[0];
  for (var i = 0; i < header.length; i++) {
    if (String(header[i]).trim().toLowerCase() === title.toLowerCase()) return i;
  }
  var idx = header.length;
  up.getRange(1, idx + 1).setValue(title);
  SpreadsheetApp.flush();
  return idx;
}

/** '1 Year' -> 12, '6 Months' -> 6, '1 Month' -> 1. 0 if unparseable. */
function parseDuration_(text) {
  var m = String(text || '').match(/(\d+)\s*(year|month)/i);
  if (!m) return 0;
  var n = parseInt(m[1], 10);
  return n && /^year/i.test(m[2]) ? n * 12 : n;
}

/** Calendar-month advance with end-of-month clamping (31 Jan + 1M -> 28 Feb). */
function addMonths_(d, n) {
  var day = d.getDate();
  var result = new Date(d.getFullYear(), d.getMonth() + n, 1);
  var daysInMonth = new Date(result.getFullYear(), result.getMonth() + 1, 0).getDate();
  result.setDate(Math.min(day, daysInMonth));
  return result;
}

function round2_(n) { return Math.round((parseFloat(n) || 0) * 100) / 100; }

/** dd/mm/yyyy (4-digit year, as used in "Bill Period"). */
function fmtDMY_(d) { return pad_(d.getDate(), 2) + '/' + pad_(d.getMonth() + 1, 2) + '/' + d.getFullYear(); }

/** dd/mm/yy (invoice-date text style used in GST RECORDS). */
function fmtShort_(d) { return pad_(d.getDate(), 2) + '/' + pad_(d.getMonth() + 1, 2) + '/' + pad_(d.getFullYear() % 100, 2); }

/** Indian-system number to words, e.g. 7080 -> "SEVEN THOUSAND EIGHTY". */
function rupeeWords_(n) {
  n = Math.round(parseFloat(n) || 0);
  if (n === 0) return 'ZERO';
  var ones = ['', 'ONE', 'TWO', 'THREE', 'FOUR', 'FIVE', 'SIX', 'SEVEN', 'EIGHT', 'NINE',
    'TEN', 'ELEVEN', 'TWELVE', 'THIRTEEN', 'FOURTEEN', 'FIFTEEN', 'SIXTEEN', 'SEVENTEEN',
    'EIGHTEEN', 'NINETEEN'];
  var tens = ['', '', 'TWENTY', 'THIRTY', 'FORTY', 'FIFTY', 'SIXTY', 'SEVENTY', 'EIGHTY', 'NINETY'];
  function two(num) {
    if (num < 20) return ones[num];
    return tens[Math.floor(num / 10)] + (num % 10 ? ' ' + ones[num % 10] : '');
  }
  function three(num) {
    var h = Math.floor(num / 100), rest = num % 100, out = [];
    if (h) out.push(ones[h] + ' HUNDRED');
    if (rest) out.push(two(rest));
    return out.join(' ');
  }
  var parts = [];
  var crore = Math.floor(n / 10000000); n %= 10000000;
  var lakh = Math.floor(n / 100000); n %= 100000;
  var thousand = Math.floor(n / 1000); n %= 1000;
  var rest = n;
  if (crore) parts.push(three(crore) + ' CRORE');
  if (lakh) parts.push(three(lakh) + ' LAKH');
  if (thousand) parts.push(three(thousand) + ' THOUSAND');
  if (rest) parts.push(three(rest));
  return parts.join(' ').replace(/\s+/g, ' ').trim();
}
