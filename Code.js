/**
 * Code.gs — GST bill automation, LOCAL-ONLY edition (Web App).
 *
 * NEW FLOW (no tunnel, no daily trigger, no server running on the PC):
 *   You run  python run_billing.py  on the PC whenever you want to bill.
 *   That script calls THIS web app twice:
 *
 *     1. action "get_due"
 *        Reads UPCOMING BILLS. A row is DUE when Expire Date <= a cutoff
 *        (sent by Python; default = yesterday) and Bill Status is empty
 *        (or PROCESSING when --resume). For each due row it:
 *          - finds the customer's LAST invoice in GST RECORDS (fuzzy match)
 *            so its GSTIN/description/amounts/payment mode can be copied,
 *          - computes the next GLOBAL invoice no from the last row
 *            (ST/046/2026-27 -> ST/047/2026-27; restarts ST/001/2027-28
 *            after each 31 March),
 *          - marks the row PROCESSING so two runs can never overlap,
 *          - returns the list to Python.
 *
 *     2. action "complete"
 *        Python edited the .docx files and made the PDFs, then sends back
 *        per-row results. For each success this app appends the new row to
 *        GST RECORDS (Payment Status "Pending", date as text) — ONLY if
 *        that invoice number is not already there (crash-safe) — and marks
 *        Bill Status "BILLED <no> <date>". Errors are marked "ERROR ...".
 *
 * ONE-TIME SETUP:
 *   1. Put your secret here:  SECRET = '<value of appscript_secret in .env>'
 *   2. Deploy > New deployment > type: Web app
 *        Execute as: Me       Who has access: Anyone
 *      Copy the Web App URL (ends in /exec) into .env:
 *        appscript_url :- https://script.google.com/macros/s/.../exec
 *   3. In the editor run dryRunOnce() (authorize when asked) — preview only.
 *   4. On the PC:  python run_billing.py --dry-run   then without --dry-run.
 *
 * NOTE: after editing this code you must create a NEW DEPLOYMENT VERSION
 *       (Deploy > Manage deployments > Edit > Version: New) or the old
 *       code keeps running.
 */

var SHEET_ID = '1taT1_t4GITmjndE8K9ZXDrrLh9BZGD3Os-AAgniZcAM';
var SECRET = '2917f7027dd1a18a30d27bf36f59a1b3';
var UPCOMING_SHEET = 'UPCOMING BILLS';
var GST_SHEET = 'GST RECORDS';

var COL_BILL_FILE = 'Bill File';     // UPCOMING BILLS: optional exact docx name
var COL_STATUS = 'Bill Status';      // UPCOMING BILLS: PROCESSING/BILLED/ERROR
var TAG_PROCESSING = 'PROCESSING';
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
  if (String(body.secret || '') !== SECRET)
    return jsonOut_({ ok: false, error: 'bad or missing secret' });
  try {
    if (body.action === 'get_due') return jsonOut_(getDue_(body));
    if (body.action === 'complete') return jsonOut_(complete_(body));
    return jsonOut_({ ok: false, error: 'unknown action: ' + body.action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  }
}

/** Browser test that the web app is alive. */
function doGet() {
  return jsonOut_({ ok: true, service: 'gst-bill-automation',
                    actions: ['get_due', 'complete'] });
}

/** Editor-only preview: logs exactly what get_due would return (marks nothing). */
function dryRunOnce() {
  var r = getDue_({ dry: true, resume: true, secret: SECRET });
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
  var iFile = ensureColumn_(up, COL_BILL_FILE);
  var iStatus = ensureColumn_(up, COL_STATUS);

  var gstVals = gst.getDataRange().getValues();
  var due = [], claim = [];
  for (var r = 1; r < upVals.length; r++) {
    var name = String(upVals[r][iName] || '').trim();
    var expire = asDate_(upVals[r][iExp]);
    if (!name || !expire) continue;
    if (expire > cutoff) continue;
    var status = String(upVals[r][iStatus] || '').trim();
    var isFree = status === '';
    var isProcessing = status.indexOf(TAG_PROCESSING) === 0;
    if (isFree) claim.push(r);
    else if (!isProcessing || !body.resume) continue;   // BILLED/ERROR stays untouched

    var last = findLastGstRecord_(gstVals, name);
    due.push({
      key: 'up_r' + (r + 1) + '_' + fmt_(expire),
      row: r + 1,
      customer: name,
      expire: fmt_(expire),
      bill_file: String(upVals[r][iFile] || '').trim(),
      status: status,
      new_invoice_no: nextInvoiceNo_(gstVals, today),
      last_invoice_no: last ? last.invoiceNo : null,
      has_gst_record: !!last
    });
  }

  if (!body.dry && claim.length) {
    claim.forEach(function (r) {
      up.getRange(r + 1, iStatus + 1).setValue(TAG_PROCESSING + ' ' + fmt_(today));
    });
    SpreadsheetApp.flush();
  }
  return { ok: true, dry: !!body.dry, cutoff: fmt_(cutoff), claimed: claim.length, due: due };
}

/* -------------------------------------------------------------- action 2 */

function complete_(body) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var up = ss.getSheetByName(UPCOMING_SHEET);
  var gst = ss.getSheetByName(GST_SHEET);
  var today = startOfDay_(new Date());
  var upVals = up.getDataRange().getValues();
  var iName = colOf_(upVals[0], 'Customer name');
  var iStatus = colOf_(upVals[0], COL_STATUS);
  var gCols = {};
  var gstVals = gst.getDataRange().getValues();
  GST_COLUMNS.forEach(function (c) { gCols[c] = colOf_(gstVals[0], c); });

  var existing = {};                       // invoice numbers already in GST RECORDS
  for (var r = 1; r < gstVals.length; r++) {
    var v = String(gstVals[r][0] || '').trim();
    if (INVOICE_RE.test(v)) existing[v] = true;
  }

  var log = [];
  var results = body.results || [];
  for (var i = 0; i < results.length; i++) {
    var res = results[i];
    var m = String(res.key || '').match(/^up_r(\d+)_/);
    if (!m) { log.push('bad key ' + res.key); continue; }
    var row = parseInt(m[1], 10) - 1;      // 0-based data index in UPCOMING BILLS
    if (row < 1 || row >= upVals.length) { log.push('row out of range for ' + res.key); continue; }
    var name = String(upVals[row][iName] || '').trim();

    if (!res.ok) {
      var cur = iStatus >= 0 ? String(upVals[row][iStatus] || '') : '';
      if (cur.indexOf('BILLED') !== 0)       // never overwrite a real success
        setStatus_(up, row, iStatus, 'ERROR ' + fmt_(today) + ': ' + String(res.error || 'unknown').slice(0, 90));
      log.push(name + ': ERROR ' + String(res.error || '').slice(0, 80));
      continue;
    }

    var invNo = String(res.new_invoice_no || '').trim();
    if (!existing[invNo]) {
      var last = findLastGstRecord_(gstVals, name);
      if (!last) {
        setStatus_(up, row, iStatus, 'ERROR ' + fmt_(today) + ': no GST record to copy for ' + name + ' — fill manually');
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
    setStatus_(up, row, iStatus, 'BILLED ' + invNo + ' on ' + fmt_(today) +
      (res.pdf_error ? ' (PDF FAILED)' : ''));
    log.push(name + ': ' + (existing[invNo] && res.duplicate ? 'already recorded' : 'record appended') +
      ' — BILLED ' + invNo);
  }
  return { ok: true, log: log };
}

/* ------------------------------------------------------- invoice numbering */

var lastInvoiceCache_ = null;

function nextInvoiceNo_(gstVals, billDate) {
  if (lastInvoiceCache_) {                                  // sequential within one request
    var fy = fyOf_(billDate);
    var n = (fy === lastInvoiceCache_.fy) ? lastInvoiceCache_.num + 1 : 1;
    return PREFIX + '/' + pad_(n, lastInvoiceCache_.width || 3) + '/' + fy;
  }
  var last = null;
  for (var r = 1; r < gstVals.length; r++) {
    var v = String(gstVals[r][0] || '').trim();
    if (INVOICE_RE.test(v)) last = v;
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

function setStatus_(sheet, row, iStatus, text) {
  if (iStatus < 0) return;
  sheet.getRange(row + 1, iStatus + 1).setValue(text);
  SpreadsheetApp.flush();
}

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

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
