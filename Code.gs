/**
 * Code.gs — GST bill automation for the "Copy of GST_BILL_RECORDS" sheet.
 *
 * WHAT IT DOES (run daily by a time-driven trigger):
 *   1. Reads the "UPCOMING BILLS" tab. A row is DUE when its Expire Date is
 *      on/before today (minus GRACE_DAYS) and its "Bill Status" column is
 *      still empty.
 *   2. For each due customer it looks up the customer's LAST invoice in
 *      "GST RECORDS" (fuzzy name match) to copy GSTIN, Service Description,
 *      Amount, CGST, SGST, Total and Payment Mode onto the renewal.
 *   3. Computes the next global invoice number from the last row of
 *      GST RECORDS:  ST/046/2026-27 -> ST/047/2026-27.
 *      After 31 March the fiscal year rolls over and numbering restarts:
 *      ST/001/2027-28, ST/002/2027-28, ...
 *   4. POSTs {customer, bill_file?, new_invoice_no, key} to the local
 *      update_bills.py server, which edits the bill .docx (Invoice Date =
 *      today, Invoice No = the exact number above, Bill Period = advanced by
 *      the bill's OWN duration in months) and creates the PDF.
 *   5. Appends the new record row to "GST RECORDS" (Payment Status =
 *      "Pending", Payment Date/GST PAID empty — fill after collection).
 *   6. Writes "BILLED <invoice no> <date>" into the UPCOMING BILLS row's
 *      Bill Status column so the row is never processed twice.
 *
 * The Python server dedupes by `key`, so an Apps Script retry can never
 * double-bump the invoice number or shift the bill period twice.
 *
 * SETUP (one time):
 *   Extensions > Apps Script -> paste this file -> Save.
 *   1. Run dryRunOnce() first (authorize when asked) — it only LOGS what it
 *      would do, touching nothing.
 *   2. Run processDueBills() once for real (or wait for the trigger).
 *   3. Triggers > Add Trigger: function processDueBills / time-driven /
 *      Day timer (e.g. 9am–10am). The PC running update_bills.py must be ON.
 *
 * THE PC SIDE: python update_bills.py  (listens on port 8765). If Apps Script
 * runs from Google's cloud, BILLS_API must be a publicly reachable URL
 * (ngrok / cloudflared tunnel); plain LAN IPs are NOT reachable from Apps
 * Script. If it fails, the row is left unbilled and retried next run.
 */

var SHEET_ID = '1rqMS2tRen54dvgt_21BwhaQFXrqlQka8VRUaFXypnBs';
var BILLS_API = 'http://127.0.0.1:8765/update';   // change to your ngrok URL if remote
var UPCOMING_SHEET = 'UPCOMING BILLS';
var GST_SHEET = 'GST RECORDS';

var COL_BILL_FILE = 'Bill File';    // UPCOMING BILLS col E (optional): exact docx name
var COL_STATUS = 'Bill Status';     // UPCOMING BILLS col F: filled by this script
var GRACE_DAYS = 0;                 // bill this many days BEFORE the expire date (0 = on the day)
var PREFIX = 'ST';                  // invoice prefix

var INVOICE_RE = /^ST\/(\d+)\/(\d{4})-(\d{1,2})$/;
var GST_COLUMNS = ['Invoice No', 'Invoice Date', 'Customer Name', 'Customer GSTIN',
  'Service Description', 'Amount ₹', 'CGST 9% (₹)', 'SGST 9% (₹)', 'Total Amount (₹)',
  'Payment Mode', 'Payment Status', 'Payment Date', 'GST PAID'];
var STOPWORDS = ['PVT', 'LTD', 'LLP', 'PRIVATE', 'LIMITED', 'CO', 'THE', 'AND',
  'BILL', 'SOLUTIONS', 'ENTERPRISES', 'INDIA', 'STORES'];

/** Trigger entry point. */
function processDueBills() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { log_('another run is already in progress'); return; }
  try {
    run_(false);
  } finally {
    lock.releaseLock();
  }
}

/** Safe test: logs every decision, sends NOTHING, writes NOTHING. */
function dryRunOnce() {
  run_(true);
}

function run_(dry) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var up = ss.getSheetByName(UPCOMING_SHEET);
  var gst = ss.getSheetByName(GST_SHEET);
  if (!up || !gst) throw new Error('tabs ' + UPCOMING_SHEET + ' / ' + GST_SHEET + ' not found');

  var today = startOfDay_(new Date());
  var upVals = up.getDataRange().getValues();
  var upHeader = upVals[0];
  var iName = colOf_(upHeader, 'Customer name');
  var iExp = colOf_(upHeader, 'Expire Date');
  if (iName < 0 || iExp < 0) throw new Error('UPCOMING BILLS needs "Customer name" and "Expire Date" columns');
  var iFile = ensureColumn_(up, COL_BILL_FILE);
  var iStatus = ensureColumn_(up, COL_STATUS);

  var gstVals = gst.getDataRange().getValues();
  var gstHeader = gstVals[0];
  var gCols = {};
  GST_COLUMNS.forEach(function (c) { gCols[c] = colOf_(gstHeader, c); });

  var results = [];
  for (var r = 1; r < upVals.length; r++) {
    var row = upVals[r];
    var name = String(row[iName] || '').trim();
    var expireRaw = row[iExp];
    var expire = asDate_(expireRaw);
    var status = String(row[iStatus] || '').trim();
    if (!name || !expire) continue;                       // empty row
    if (status) continue;                                 // already handled

    var due = addDays_(expire, -GRACE_DAYS) <= today;
    if (!due) { results.push(name + ': not due (expires ' + fmt_(expire) + ')'); continue; }

    var billFile = String(row[iFile] || '').trim() || null;
    try {
      var lastRec = findLastGstRecord_(gstVals, name);    // may be null
      var invNo = nextInvoiceNo_(gstVals, today);
      var key = 'up_r' + (r + 1) + '_' + fmt_(expire);
      var rec = {
        customer: name,
        new_invoice_no: invNo,
        period_shift: 'auto',               // advance by the bill's own duration (1mo -> 1mo, 3mo -> 3mo ...)
        key: key
      };
      if (billFile) rec.bill_file = billFile;
      if (lastRec && lastRec.invoiceNo) rec.invoice_no = lastRec.invoiceNo;

      if (dry) {
        results.push('[dry-run] ' + name + ' -> would bill ' + invNo +
          (lastRec ? ' (copying ' + lastRec.invoiceNo + ')' : ' (WARNING: no GST record match — row would be left for manual fill)'));
        continue;
      }

      var resp = callPython_(rec);
      if (!resp.ok) throw new Error(resp.error || 'python returned ok=false');
      if (resp.new_invoice_no && resp.new_invoice_no !== invNo)
        throw new Error('python wrote ' + resp.new_invoice_no + ' but sheet says ' + invNo);

      if (resp.duplicate) {
        // python already made this bill (retry) — do NOT append a second record
        setStatus_(up, r, iStatus, 'BILLED ' + invNo + ' on ' + fmt_(today) + ' (dedup)');
        results.push(name + ': python reports duplicate for ' + invNo + ' — status marked, no new record');
        continue;
      }

      if (!lastRec) throw new Error('no matching GST RECORDS row for ' + name + ' — record NOT added, fill manually');

      appendGstRecord_(gst, gCols, {
        invoiceNo: invNo,
        invoiceDate: String(resp.invoice_date || '').replace(/-/g, '/'),
        customer: lastRec.customer,                        // sheet's own spelling
        gstin: lastRec.gstin, desc: lastRec.desc,
        amount: lastRec.amount, cgst: lastRec.cgst, sgst: lastRec.sgst, total: lastRec.total,
        mode: lastRec.mode
      });
      setStatus_(up, r, iStatus, 'BILLED ' + invNo + ' on ' + fmt_(today));
      results.push(name + ': BILLED ' + invNo + ' (was ' + lastRec.invoiceNo + ')');
      // keep local numbering in sync for further rows in this run
      lastInvoiceCache_ = { num: parseInt(invNo.match(INVOICE_RE)[1], 10), fy: fyOf_(today) };
    } catch (e) {
      results.push('ERROR ' + name + ': ' + e.message);
      if (!dry) setStatus_(up, r, iStatus, 'ERROR ' + fmt_(today) + ': ' + String(e.message).slice(0, 100));
    }
  }
  log_('--- run summary ---\n' + results.join('\n'));
  if (!dry) ss.toast(results.join('\n').slice(0, 250), 'Bill run', 10);
}

/* ------------------------------------------------------------------ python */

function callPython_(rec) {
  var resp = UrlFetchApp.fetch(BILLS_API, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(rec),
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  var body = resp.getContentText() || '{}';
  var parsed;
  try { parsed = JSON.parse(body); } catch (e) { parsed = { ok: false, error: 'non-JSON: ' + body.slice(0, 120) }; }
  if (code !== 200 && !parsed.ok) parsed = { ok: false, error: 'HTTP ' + code + ': ' + String(parsed.error || body).slice(0, 150) };
  return parsed;
}

/* ------------------------------------------------------- invoice numbering */

var lastInvoiceCache_ = null;

function nextInvoiceNo_(gstVals, billDate) {
  if (lastInvoiceCache_) {                                  // sequential within one run
    var fy = fyOf_(billDate);
    var n = (fy === lastInvoiceCache_.fy) ? lastInvoiceCache_.num + 1 : 1;
    return PREFIX + '/' + pad_(n, lastInvoiceCache_.width || 3) + '/' + fy;
  }
  var last = null, lastRow = -1;
  for (var r = 1; r < gstVals.length; r++) {
    var v = String(gstVals[r][0] || '').trim();
    if (INVOICE_RE.test(v)) { last = v; lastRow = r; }
  }
  if (!last) return PREFIX + '/001/' + fyOf_(billDate);
  var m = last.match(INVOICE_RE);
  var lastNum = parseInt(m[1], 10), lastFy = m[2] + '-' + m[3];
  var fyNow = fyOf_(billDate);
  var num = (fyNow === lastFy) ? lastNum + 1 : 1;
  lastInvoiceCache_ = { num: num, fy: fyNow, width: m[1].length };
  return PREFIX + '/' + pad_(num, m[1].length) + '/' + fyNow;
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

function ensureColumn_(sheet, title) {
  var lastCol = sheet.getLastColumn();                      // ignores trailing empty cols
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = header.indexOf(title);
  if (idx >= 0) return idx;
  sheet.getRange(1, lastCol + 1).setValue(title);           // append new column
  SpreadsheetApp.flush();
  return lastCol;                                           // 0-based index of the new column
}

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
function log_(msg) { Logger.log(msg); }
