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
    if (body.action === 'debug_rebuild') return jsonOut_(debugRebuild_(body));
    // ---- customer-doc billing (edit-in-place + Drive archive) actions ----
    if (body.action === 'sync_docs') return jsonOut_(syncDocs_(body));
    if (body.action === 'check_docs') return jsonOut_(checkDocs_(body));
    if (body.action === 'rebuild_doc') return jsonOut_(rebuildDoc_(body));
    if (body.action === 'list_pdfs') return jsonOut_(listPdfs_(body));
    if (body.action === 'get_pdf') return jsonOut_(getPdf_(body));
    if (body.action === 'map_doc') return jsonOut_(mapDoc_(body));
    if (body.action === 'add_customer') return jsonOut_(addCustomer_(body));
    if (body.action === 'repair_doc') return jsonOut_(repairDoc_(body));
    if (body.action === 'update_payment') return jsonOut_(updatePayment_(body));
    if (body.action === 'refresh_pdf') return jsonOut_(refreshPdf_(body));
    if (body.action === 'bill_prefill') return jsonOut_(billPrefill_(body));
    if (body.action === 'backfill_periods') return jsonOut_(backfillPeriods_(body));
    if (body.action === 'bill_edit_prefill') return jsonOut_(billEditPrefill_(body));
    if (body.action === 'edit_bill') return jsonOut_(editBill_(body));
    if (body.action === 'delete_customer') return jsonOut_(deleteCustomer_(body));
    if (body.action === 'list_deleted') return jsonOut_(listDeletedCustomers_(body));
    if (body.action === 'restore_customer') return jsonOut_(restoreDeletedCustomer_(body));
    if (body.action === 'purge_deleted_customer') return jsonOut_(purgeDeletedCustomer_(body));
    return jsonOut_({ ok: false, error: 'unknown action: ' + body.action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  }
}

/** Browser test that the web app is alive. */
function doGet() {
  return jsonOut_({ ok: true, service: 'gst-bill-automation v6 (pure DocumentApp)',
                    actions: ['get_due', 'complete', 'set_props',
                              'ping', 'list_all', 'bill_customer', 'make_pdf',
                              'sync_docs', 'check_docs', 'rebuild_doc',
                              'list_pdfs', 'get_pdf', 'map_doc', 'add_customer',
                    'repair_doc', 'update_payment', 'refresh_pdf',
                    'bill_prefill', 'bill_edit_prefill', 'edit_bill', 'backfill_periods',
                    'delete_customer', 'list_deleted', 'restore_customer', 'purge_deleted_customer',
                              'debug_rebuild'] });
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
  var allowed = { SECRET: 1, SHEET_ID: 1, BILL_TEMPLATE_DOC_ID: 1,
                  DOCS_FOLDER_ID: 1, PDFS_FOLDER_ID: 1, NEW_CUSTOMER_TEMPLATE_ID: 1 };
  var set = [];
  var props = body.props || {};

  // Validate the new-customer template before storing it as the active template.
  if (Object.prototype.hasOwnProperty.call(props, 'NEW_CUSTOMER_TEMPLATE_ID')) {
    var candidateTemplateId = String(props.NEW_CUSTOMER_TEMPLATE_ID || '').trim();
    if (!candidateTemplateId) throw new Error('NEW_CUSTOMER_TEMPLATE_ID cannot be blank.');
    validateNewCustomerTemplate_(candidateTemplateId);
  }

  Object.keys(props).forEach(function (k) {
    if (!allowed[k]) return;
    var v = String(props[k] || '').trim();
    if (!v) return;
    PROPS.setProperty(k, v);
    set.push(k);
  });
  if (!set.length) throw new Error('no valid properties to set (allowed: SECRET, SHEET_ID, BILL_TEMPLATE_DOC_ID, DOCS_FOLDER_ID, PDFS_FOLDER_ID, NEW_CUSTOMER_TEMPLATE_ID)');
  if (set.indexOf('SECRET') >= 0) SECRET = PROPS.getProperty('SECRET');
  if (set.indexOf('SHEET_ID') >= 0) SHEET_ID = PROPS.getProperty('SHEET_ID');
  return { ok: true, set: set, template_validated: set.indexOf('NEW_CUSTOMER_TEMPLATE_ID') >= 0 };
}


/**
 * Checks that the configured onboarding source is a readable native Google Doc
 * containing the full placeholder contract required by addCustomerLocked_.
 */
function validateNewCustomerTemplate_(templateId) {
  var file;
  try {
    file = DriveApp.getFileById(String(templateId || '').trim());
  } catch (e) {
    throw new Error('Cannot access the master template. Check the Google Doc ID and share the document with the Apps Script owner.');
  }
  if (file.getMimeType() !== MimeType.GOOGLE_DOCS) {
    throw new Error('The master template must be a native Google Doc. Upload the DOCX, open it with Google Docs, then choose File → Save as Google Docs.');
  }

  var doc;
  try {
    doc = DocumentApp.openById(file.getId());
  } catch (e2) {
    throw new Error('The master template file is not readable as a Google Doc by the Apps Script owner.');
  }
  var docText = doc.getBody().getText();
  var required = [
    'BUYER_NAME', 'BUYER_ADDRESS', 'BUYER_GSTIN_LINE',
    'BUYER_STATE', 'BUYER_STATE_CODE', 'INVOICE_DATE', 'INVOICE_NO',
    'USER_ID', 'PLACE_OF_SUPPLY', 'PLACE_STATE_CODE', 'BILL_PERIOD',
    'PLAN_LINE_1', 'PLAN_LINE_2', 'CHARGES', 'LINE_AMOUNT',
    'BASIC_AMOUNT', 'CGST_AMOUNT', 'SGST_AMOUNT', 'ROUND_OFF',
    'TOTAL_AMOUNT', 'AMOUNT_IN_WORDS'
  ];
  var missing = required.filter(function (key) {
    return docText.indexOf('[[' + key + ']]') < 0;
  });
  if (missing.length) {
    throw new Error('Master template is missing required placeholders: ' +
      missing.map(function (key) { return '[[' + key + ']]'; }).join(', '));
  }
  return { valid: true, documentId: file.getId(), placeholders: required.length };
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
      var newPRow = appendGstRecord_(gst, gCols, {
        invoiceNo: invNo,
        invoiceDate: String(res.invoice_date || '').replace(/-/g, '/'),
        customer: last.customer, gstin: last.gstin, desc: last.desc,
        amount: last.amount, cgst: last.cgst, sgst: last.sgst, total: last.total,
        mode: last.mode
      });
      var pExp = asDate_(res.new_expire);
      var pMon = parseInt(res.duration_months, 10);
      if (pExp && pMon > 0)
        writeBillPeriod_(gst, ensureBillPeriodCol_(gst), newPRow - 1,
          addDays_(addMonths_(pExp, -pMon), 1), pExp);
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

/**
 * Heal legacy/text expiry values and enforce a two-digit day/month display.
 * Runs on list refresh and after sorting, so existing rows are corrected too.
 */
function normalizeUpcomingExpireDates_(up, iExp) {
  var lastRow = up.getLastRow();
  if (lastRow < 2 || iExp < 0) return;
  var range = up.getRange(2, iExp + 1, lastRow - 1, 1);
  var values = range.getValues();
  var changed = false;
  for (var i = 0; i < values.length; i++) {
    var value = values[i][0];
    if (value instanceof Date || value === '' || value === null) continue;
    var parsed = asDate_(value);
    if (parsed) {
      values[i][0] = parsed;
      changed = true;
    }
  }
  if (changed) range.setValues(values);
  range.setNumberFormat('dd/mm/yy');
}

/** Re-orders the data rows of UPCOMING BILLS by Expire Date, nearest first. */
/** Find the UPCOMING BILLS serial-number column without creating a new column. */
function upcomingSrNoColumn_(headers) {
  for (var i = 0; i < headers.length; i++) {
    var key = String(headers[i] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    if (key === 'srno' || key === 'sno' || key === 'serialno' ||
        key === 'serialnumber' || key === 'srnumber') return i;
  }
  return -1;
}

/** Return one greater than the highest existing numeric Sr No. */
function nextUpcomingSrNo_(up, srNoColumn) {
  if (srNoColumn < 0) throw new Error('UPCOMING BILLS has no Sr No column (accepted headers: Sr No, Sr. No., S.No., Serial No)');
  var values = up.getDataRange().getValues();
  var maxNumber = 0;
  for (var r = 1; r < values.length; r++) {
    var raw = String(values[r][srNoColumn] === null || values[r][srNoColumn] === undefined
      ? '' : values[r][srNoColumn]).trim();
    if (!raw) continue;
    var number = /^\d+$/.test(raw) ? parseInt(raw, 10) : 0;
    if (number > maxNumber) maxNumber = number;
  }
  return maxNumber + 1;
}

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
  normalizeUpcomingExpireDates_(up, iExp);
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
      best = { invoiceNo: inv, invoiceDate: d ? fmtShort_(d) : String(row[1] || '').trim(),
               customer: String(row[2]).trim(), gstin: gstinDocumentValue_(row[3]),
               desc: String(row[4] || '').trim(), amount: row[5], cgst: row[6], sgst: row[7],
               total: row[8], mode: String(row[9] || '').trim() };
      bestScore = s; bestDate = d;
    }
  }
  return best;
}

/** Normalize optional GSTIN for invoice/document operations. */
function gstinDocumentValue_(value) {
  var text = String(value === null || value === undefined ? '' : value).trim();
  return !text || text === '-' ? '' : text;
}

/** GST RECORDS uses a visible hyphen instead of a blank optional GSTIN. */
function gstinSheetValue_(value) {
  return gstinDocumentValue_(value) || '-';
}

function appendGstRecord_(gst, gCols, v) {
  var rowIdx = gst.getLastRow() + 1;
  var row = new Array(GST_COLUMNS.length).fill('');
  row[gCols['Invoice No']] = v.invoiceNo;
  row[gCols['Invoice Date']] = v.invoiceDate;
  row[gCols['Customer Name']] = v.customer;
  row[gCols['Customer GSTIN']] = gstinSheetValue_(v.gstin);
  // Keep full wording on the Google Doc/PDF; store a compact label in GST RECORDS.
  row[gCols['Service Description']] = conciseServiceDescription_(v.desc) || v.desc;
  row[gCols['Amount ₹']] = v.amount;
  row[gCols['CGST 9% (₹)']] = v.cgst;
  row[gCols['SGST 9% (₹)']] = v.sgst;
  row[gCols['Total Amount (₹)']] = v.total;
  row[gCols['Payment Mode']] = v.mode;
  row[gCols['Payment Status']] = v.payment_status || 'Pending';
  var range = gst.getRange(rowIdx, 1, 1, GST_COLUMNS.length);
  range.setValues([row]);
  var dCell = gst.getRange(rowIdx, gCols['Invoice Date'] + 1);
  dCell.setNumberFormat('@');                               // keep dd/mm/yy as text
  dCell.setValue(v.invoiceDate);
  SpreadsheetApp.flush();
  return rowIdx;                                            // 1-based sheet row
}

/** Hidden "Bill Period" column in GST RECORDS: stores 'dd/mm/yyyy to
 *  dd/mm/yyyy' per row — a track of every bill's period, older bills
 *  included (the bill form's period prefill reads it). Created on first use
 *  and hidden (users can unhide in the sheet; it will re-hide on writes). */
function ensureBillPeriodCol_(gst) {
  var last = gst.getLastColumn();
  var header = gst.getRange(1, 1, 1, last).getValues()[0];
  for (var i = 0; i < header.length; i++) {
    if (String(header[i]).trim() === 'Bill Period') {
      try { gst.hideColumns(i + 1); } catch (e) {}
      return i;                                             // 0-based
    }
  }
  var col = last + 1;                                       // 1-based new column
  gst.getRange(1, col).setValue('Bill Period');
  try { gst.hideColumns(col); } catch (e) {}
  return col - 1;
}

/** Write one row's bill period (start/end Dates) into the hidden column. */
function writeBillPeriod_(gst, c0, r0, ps, pe) {
  if (c0 === null || c0 < 0 || !(r0 >= 0) || !ps || !pe) return;
  gst.getRange(r0 + 1, c0 + 1).setValue(fmtDMY_(ps) + ' to ' + fmtDMY_(pe));
}

var customerNameAliasCache_ = null;

function loadCustomerNameAliases_() {
  if (customerNameAliasCache_) return customerNameAliasCache_;
  try { customerNameAliasCache_ = JSON.parse(PROPS.getProperty('CUSTOMER_NAME_ALIASES') || '{}'); }
  catch (e) { customerNameAliasCache_ = {}; }
  return customerNameAliasCache_;
}

function canonicalCustomerName_(value) {
  var key = normName_(value);
  var aliases = loadCustomerNameAliases_();
  return aliases[key] ? normName_(aliases[key]) : key;
}

function nameScore_(a, b) {
  var na = normName_(a), nb = normName_(b);
  if (!na || !nb) return 0;
  var ca = canonicalCustomerName_(a), cb = canonicalCustomerName_(b);
  if (ca && cb && ca === cb) return 1;
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

/** Date-only value for invoice operations, based on the India business timezone. */
function indiaToday_() {
  var iso = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
  return parseISO_(iso);
}
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
  normalizeUpcomingExpireDates_(up, iExp);
  upVals = up.getDataRange().getValues();
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
  var custs = [];
  for (var r = 1; r < upVals.length; r++) {
    var name = String(upVals[r][iName] || '').trim();
    if (!name) continue;
    custs.push({ row: r + 1, name: name, invoices: [] });
  }
  // Each gst row belongs to its BEST-matching customer ONLY. Plain fuzzy
  // matching would hand "Dilip K Jain & Associates 1" and "…2" each other's
  // bills — the names differ by one character, similarity ~0.95.
  // A row only attaches on an EXACT name match (>= 0.95): recurring rows are
  // copied from the UPCOMING sheet so they match exactly; anything else
  // (one-time bills like "Ornellas High School Switch", "Spes Manning
  // Solutions llp Router") becomes its own entry instead of bleeding into a
  // similarly-named customer.
  var attached = {};
  for (var k = 0; k < gstRows.length; k++) {
    var best = -1, bestScore = 0;
    for (var c2 = 0; c2 < custs.length; c2++) {
      var s = nameScore_(custs[c2].name, gstRows[k].customer);
      if (s > bestScore) { bestScore = s; best = c2; }
    }
    if (best >= 0 && bestScore >= 0.95) { custs[best].invoices.push(gstRows[k]); attached[k] = true; }
  }
  // GST rows that matched no UPCOMING customer -> their own customers
  // (one-time bills: doc + invoice exist, but there is no renewal row)
  var virt = {}, vOrder = [];
  for (var k2 = 0; k2 < gstRows.length; k2++) {
    if (attached[k2]) continue;
    var nm2 = gstRows[k2].customer || 'Unknown';
    if (!virt[nm2]) { virt[nm2] = []; vOrder.push(nm2); }
    virt[nm2].push(gstRows[k2]);
  }
  for (var v1 = 0; v1 < vOrder.length; v1++) {
    var inv2 = virt[vOrder[v1]], last2 = inv2[0];
    customers.push({
      row: -1, name: vOrder[v1], expire: '', expire_iso: '',
      duration: '', bill_file: '', address: '', userid: '',
      gstin: last2 ? last2.gstin : '',
      last_invoice: last2 ? stripRow_(last2) : null,
      invoices: inv2.map(stripRow_)
    });
  }
  for (var c3 = 0; c3 < custs.length; c3++) {
    var r2 = custs[c3].row - 1;
    var name3 = custs[c3].name;
    var expire = asDate_(upVals[r2][iExp]);
    var invoices = custs[c3].invoices;
    var lastInv = invoices.length ? invoices[0] : null;

    customers.push({
      row: custs[c3].row,
      name: name3,
      expire: expire ? fmtDMY_(expire) : '',
      expire_iso: expire ? iso_(expire) : '',
      duration: iDur >= 0 ? String(upVals[r2][iDur] || '').trim() : '',
      bill_file: iFile >= 0 ? String(upVals[r2][iFile] || '').trim() : '',
      address: iAddr >= 0 ? String(upVals[r2][iAddr] || '').trim() : '',
      userid: iUser >= 0 ? String(upVals[r2][iUser] || '').trim() : '',
      gstin: lastInv ? lastInv.gstin : '',
      last_invoice: lastInv ? stripRow_(lastInv) : null,
      invoices: invoices.map(stripRow_)
    });
  }
  // doc-mapped names with NEITHER a UPCOMING row NOR a GST row — so every
  // customer with a Google Doc on Drive shows up in the app
  var dmap; try { dmap = JSON.parse(PROPS.getProperty('DOCS_MAP') || '{}'); } catch (e5) { dmap = {}; }
  for (var dn in dmap) {
    var known = customers.some(function (c9) { return nameScore_(c9.name, dn) >= 0.95; });
    if (known) continue;
    customers.push({
      row: -1, name: dn, expire: '', expire_iso: '', duration: '', bill_file: '',
      address: '', userid: '', gstin: '', last_invoice: null, invoices: [],
      doc_only: true
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
  var today = indiaToday_();
  // --- invoice-date guard + early-billing override -------------------------
  // A renewal may be invoiced only AFTER the current plan expires (from the
  // next day on). To bill early, the caller must send `invoice_date`
  // (yyyy-mm-dd) strictly after the expiry — that date is then used on the
  // bill AND in GST RECORDS, so the invoice reads as if it was created after
  // the plan expired. Without the override the request is rejected.
  var invoiceDate = today;
  var ovRaw = String(body.invoice_date || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(ovRaw)) {
    var ovp = ovRaw.split('-');
    invoiceDate = new Date(Number(ovp[0]), Number(ovp[1]) - 1, Number(ovp[2]));
    if (invoiceDate <= oldExp) {
      throw new Error('override invoice date (' + fmtDMY_(invoiceDate) +
        ') must be AFTER ' + custName + '"s expiry (' + fmtDMY_(oldExp) + ')');
    }
  } else if (today <= oldExp) {
    throw new Error(custName + ' is covered until ' + fmtDMY_(oldExp) +
      ' — a bill can be invoiced only after expiry (from ' + fmtDMY_(addDays_(oldExp, 1)) +
      '). To bill early, enable the override and pick an invoice date after the expiry.');
  }
  var periodStart = addDays_(oldExp, 1);
  var periodEnd = addDays_(addMonths_(periodStart, months), -1);
  var invoiceNo = nextInvoiceNo_(gstVals, invoiceDate);

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

  // --- edit the customer's own Google Doc in place, export + archive PDF ---
  // Nothing is written to the sheet unless this succeeds.
  var bill = {
    buyer: last.customer,
    address: address,
    gstin: last.gstin,
    userid: userid,
    invoiceNo: invoiceNo,
    invoiceDate: fmtShort_(invoiceDate),
    periodStart: fmtDMY_(periodStart),
    periodEnd: fmtDMY_(periodEnd),
    desc: desc,
    amount: amount, cgst: cgst, sgst: sgst, total: total,
    mode: mode
  };
  var lastMonths = parseDuration_(iDur >= 0 ? String(upVals[row][iDur] || '') : '');
  var pdf;
  try {
    pdf = billViaCustomerDoc_(custName, last, oldExp, lastMonths, bill);
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
      var newRow = appendGstRecord_(gst, gCols, {
        invoiceNo: invoiceNo,
        invoiceDate: fmtShort_(invoiceDate),
        customer: last.customer, gstin: last.gstin, desc: desc,
        amount: amount, cgst: cgst, sgst: sgst, total: total, mode: mode
      });
      writeBillPeriod_(gst, ensureBillPeriodCol_(gst), newRow - 1,
        addDays_(addMonths_(periodEnd, -months), 1), periodEnd);
      appended = true;
    }

    var cell = up.getRange(row + 1, iExp + 1);
    cell.setNumberFormat('dd/mm/yy');
    cell.setValue(periodEnd);
    if (iDur >= 0) up.getRange(row + 1, iDur + 1).setValue(durationText_(months));

    if (address && iAddr >= 0) up.getRange(row + 1, iAddr + 1).setValue(address);
    if (userid && iUser >= 0) up.getRange(row + 1, iUser + 1).setValue(userid);

    sortUpcomingByExpire_(up, iExp);
  } catch (err) {
    sheetError = String(err && err.message || err);
  }

  return {
    ok: true,
    customer: custName,
    invoice_no: invoiceNo,
    invoice_date: fmtShort_(invoiceDate),
    appended: appended,
    sheet_updated: sheetError === null,
    sheet_error: sheetError,
    duration_months: months,
    period: { start: fmtDMY_(periodStart), end: fmtDMY_(periodEnd) },
    new_expire: fmtDMY_(periodEnd),
    amount: amount, cgst: cgst, sgst: sgst, total: total,
    desc: desc, mode: mode,
    pdf_source: pdf.source,
    pdf_file: pdf.file || '',
    pdf_folder: pdf.folderUrl || '',
    pdf_name: custName.replace(/[\\\/:*?"<>|]+/g, '_').replace(/\s+/g, '_') + '_BILL.pdf',
    pdf_base64: Utilities.base64Encode(pdfBlob.getBytes())
  };
}

/* ------------------------------------------- customer-doc billing (Drive) ---
 *
 * Each customer has ONE Google Doc (in the shared Gst_bills folder). Billing
 * edits that doc IN PLACE — the previous bill's values (reconstructed from
 * GST RECORDS + the Duration column) are replaced with the new bill's — then
 * the doc is exported to PDF and archived under BILLS/<Customer>/ on Drive.
 * The sheet is updated only after the PDF is safely archived.
 * -------------------------------------------------------------------------- */

function sanitizeFile_(s) {
  return String(s).replace(/[\\\/:*?"<>|]+/g, '_').replace(/\s+/g, '_');
}

/** The Gst_bills folder (source docs). */
function docsFolder_() {
  var id = PROPS.getProperty('DOCS_FOLDER_ID');
  if (!id) throw new Error('DOCS_FOLDER_ID not set — register the Gst_bills folder id first (test_backend.py folder)');
  return DriveApp.getFolderById(id);
}

/** Root folder for archived PDFs: prop PDFS_FOLDER_ID, or "BILLS" inside the
 *  Gst_bills folder (auto-created on first use). */
function pdfsRoot_() {
  var id = PROPS.getProperty('PDFS_FOLDER_ID');
  if (id) return DriveApp.getFolderById(id);
  var docs = docsFolder_();
  var it = docs.getFoldersByName('BILLS');
  return it.hasNext() ? it.next() : docs.createFolder('BILLS');
}

function loadCustomerFolderAliases_() {
  try { return JSON.parse(PROPS.getProperty('CUSTOMER_FOLDER_ALIASES') || '{}'); }
  catch (e) { return {}; }
}

function customerFolderName_(custName) {
  var aliases = loadCustomerFolderAliases_();
  return aliases[normName_(custName)] || sanitizeFile_(custName);
}

function customerFolder_(custName, create) {
  var root = pdfsRoot_();
  var folderName = customerFolderName_(custName);
  var it = root.getFoldersByName(folderName);
  if (it.hasNext()) return it.next();
  return create ? root.createFolder(folderName) : null;
}

/** Doc id mapped to a customer (DOCS_MAP prop, built by syncDocs_). */
function customerDocId_(custName) {
  var raw = PROPS.getProperty('DOCS_MAP');
  if (!raw) return '';
  try {
    var m = JSON.parse(raw);
    return m[custName] || '';
  } catch (e) { return ''; }
}

function saveDocsMap_(map) { PROPS.setProperty('DOCS_MAP', JSON.stringify(map)); }

/** Manual overrides: customer -> doc id set via map_doc. Survives sync_docs
 *  (which rebuilds the auto-matched map); manual entries always win. */
function loadManualMap_() {
  var raw = PROPS.getProperty('DOCS_MAP_MANUAL');
  if (!raw) return {};
  try { return JSON.parse(raw); } catch (e) { return {}; }
}
function saveManualMap_(m) { PROPS.setProperty('DOCS_MAP_MANUAL', JSON.stringify(m)); }

/**
 * Keep both the previous account names and the new buyer name resolving to
 * the same Doc and Drive archive. This prevents independent sheet renames
 * from breaking future billing or access to older PDFs.
 */
function rememberCustomerRenameAliases_(oldNames, newName, docId, canonicalGstName) {
  var names = (oldNames || []).concat([newName]).map(function (v) {
    return String(v || '').trim();
  }).filter(Boolean);
  var unique = [];
  names.forEach(function (v) {
    if (unique.indexOf(v) < 0) unique.push(v);
  });
  if (!unique.length) return;

  var root = pdfsRoot_();
  var aliases = loadCustomerFolderAliases_();
  var canonical = '';
  for (var i = 0; i < unique.length && !canonical; i++) {
    canonical = aliases[normName_(unique[i])] || '';
  }
  for (var j = 0; j < unique.length && !canonical; j++) {
    var it = root.getFoldersByName(sanitizeFile_(unique[j]));
    if (it.hasNext()) canonical = it.next().getName();
  }
  if (!canonical) canonical = sanitizeFile_(unique[0]);
  unique.forEach(function (v) { aliases[normName_(v)] = canonical; });
  PROPS.setProperty('CUSTOMER_FOLDER_ALIASES', JSON.stringify(aliases));

  // Maintain stable identity across the two independently editable sheets.
  // If UPCOMING BILLS is renamed without GST RECORDS, the app must still be
  // able to locate the old GST rows for future billing and edit operations.
  var nameAliases = {};
  try { nameAliases = JSON.parse(PROPS.getProperty('CUSTOMER_NAME_ALIASES') || '{}'); } catch (e) {}
  var canonical = String(canonicalGstName || newName || unique[0]).trim();
  var previousCanonicals = {};
  unique.forEach(function (v) {
    var existingCanonical = nameAliases[normName_(v)];
    if (existingCanonical) previousCanonicals[normName_(existingCanonical)] = true;
  });
  Object.keys(nameAliases).forEach(function (key) {
    if (previousCanonicals[normName_(nameAliases[key])]) nameAliases[key] = canonical;
  });
  unique.forEach(function (v) { nameAliases[normName_(v)] = canonical; });
  nameAliases[normName_(canonical)] = canonical;
  PROPS.setProperty('CUSTOMER_NAME_ALIASES', JSON.stringify(nameAliases));
  customerNameAliasCache_ = nameAliases;

  if (docId) {
    var map = {};
    try { map = JSON.parse(PROPS.getProperty('DOCS_MAP') || '{}'); } catch (e) {}
    unique.forEach(function (v) { map[v] = docId; });
    saveDocsMap_(map);

    var manual = loadManualMap_();
    unique.forEach(function (v) { manual[v] = docId; });
    saveManualMap_(manual);
  }
}


/* --------------------------------------------------------------- sync ----- */

/** sync_docs: walk the docs folder recursively, fuzzy-match every Google Doc
 *  to a UPCOMING BILLS customer and store the mapping. Read-only for docs. */
/* ------------------------------------------------------- customer recycle bin */

/** A dedicated recycle bin under the Gst_Bills source folder. */
function customerBinFolder_(create) {
  var root = docsFolder_();
  var it = root.getFoldersByName('BIN');
  if (it.hasNext()) return it.next();
  return create ? root.createFolder('BIN') : null;
}

function loadJsonProperty_(key) {
  try { return JSON.parse(PROPS.getProperty(key) || '{}'); }
  catch (e) { return {}; }
}

/** Resolve the exact account identity plus all recorded aliases. */
function customerIdentityContext_(customer) {
  var canonical = canonicalCustomerName_(customer);
  var keys = {};
  function add(value) {
    var key = normName_(value);
    if (key) keys[key] = true;
  }
  add(customer);
  add(canonical);

  var names = loadCustomerNameAliases_();
  Object.keys(names).forEach(function (key) {
    if (canonicalCustomerName_(key) === canonical) {
      add(key);
      add(names[key]);
    }
  });

  var docsMap = loadJsonProperty_('DOCS_MAP');
  var manualMap = loadManualMap_();
  [docsMap, manualMap].forEach(function (map) {
    Object.keys(map).forEach(function (key) {
      if (canonicalCustomerName_(key) === canonical) add(key);
    });
  });

  return {
    canonical: canonical,
    keys: keys,
    matches: function (value) {
      var key = normName_(value);
      return !!key && (!!keys[key] || canonicalCustomerName_(value) === canonical);
    }
  };
}

function packBinCell_(value) {
  if (value instanceof Date) return { __gstBinType: 'date', value: value.toISOString() };
  return value;
}
function unpackBinCell_(value) {
  if (value && typeof value === 'object' && value.__gstBinType === 'date')
    return new Date(value.value);
  return value;
}

/** Snapshot the complete source rows, including hidden columns and formulas. */
function snapshotSheetRows_(sheet, rowNumbers) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) {
    return String(v === null || v === undefined ? '' : v);
  });
  var rows = rowNumbers.map(function (rowNumber) {
    var range = sheet.getRange(rowNumber, 1, 1, lastCol);
    return {
      sourceRow: rowNumber,
      values: range.getValues()[0].map(packBinCell_),
      formulas: range.getFormulas()[0]
    };
  });
  return { headers: headers, rows: rows };
}

function binHeaderIndex_(headers, title) {
  var key = String(title || '').trim().toLowerCase();
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim().toLowerCase() === key) return i;
  }
  return -1;
}

/** Restore saved rows by header name, so a later-added sheet column is safe. */
function restoreSheetSnapshot_(sheet, snapshot, kind, identity) {
  if (!snapshot || !snapshot.rows || !snapshot.rows.length) return 0;
  var lastCol = sheet.getLastColumn();
  var currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) {
    return String(v === null || v === undefined ? '' : v);
  });
  var sourceNameIndex = binHeaderIndex_(snapshot.headers, kind === 'gst' ? 'Customer Name' : 'Customer name');
  var sourceInvoiceIndex = binHeaderIndex_(snapshot.headers, 'Invoice No');
  var currentInvoiceIndex = binHeaderIndex_(currentHeaders, 'Invoice No');
  var currentNameIndex = binHeaderIndex_(currentHeaders, kind === 'gst' ? 'Customer Name' : 'Customer name');
  var existingInvoices = {};
  var existingNames = {};
  var existingSrNumbers = {};
  var maxSrNumber = 0;
  var srColumn = kind === 'upcoming' ? upcomingSrNoColumn_(currentHeaders) : -1;
  var currentValues = sheet.getDataRange().getValues();
  for (var r = 1; r < currentValues.length; r++) {
    if (kind === 'gst' && currentInvoiceIndex >= 0) {
      existingInvoices[String(currentValues[r][currentInvoiceIndex] || '').trim()] = true;
    }
    if (kind === 'upcoming' && currentNameIndex >= 0) {
      var currentName = String(currentValues[r][currentNameIndex] || '').trim();
      if (currentName) existingNames[normName_(currentName)] = true;
      if (srColumn >= 0) {
        var currentSr = parseInt(currentValues[r][srColumn], 10);
        if (isFinite(currentSr) && currentSr > 0) {
          existingSrNumbers[String(currentSr)] = true;
          maxSrNumber = Math.max(maxSrNumber, currentSr);
        }
      }
    }
  }

  var appended = 0;
  snapshot.rows.forEach(function (saved) {
    var savedName = sourceNameIndex >= 0 ? String(unpackBinCell_(saved.values[sourceNameIndex]) || '').trim() : '';
    var savedInvoice = sourceInvoiceIndex >= 0 ? String(unpackBinCell_(saved.values[sourceInvoiceIndex]) || '').trim() : '';
    if (kind === 'gst' && savedInvoice && existingInvoices[savedInvoice]) return;
    if (kind === 'upcoming' && savedName && existingNames[normName_(savedName)]) return;

    var values = new Array(lastCol).fill('');
    var formulasToSet = [];
    for (var c = 0; c < currentHeaders.length; c++) {
      var currentKey = String(currentHeaders[c] || '').trim().toLowerCase();
      var sourceIndex = -1;
      for (var s = 0; s < snapshot.headers.length; s++) {
        if (String(snapshot.headers[s] || '').trim().toLowerCase() === currentKey && currentKey) {
          sourceIndex = s; break;
        }
      }
      // Blank/unlabelled columns retain their original positional value.
      if (sourceIndex < 0 && !currentKey && c < saved.values.length) sourceIndex = c;
      if (sourceIndex < 0 || sourceIndex >= saved.values.length) continue;
      values[c] = unpackBinCell_(saved.values[sourceIndex]);
      var formula = saved.formulas && saved.formulas[sourceIndex] || '';
      if (formula) formulasToSet.push({ column: c + 1, formula: formula });
    }
    var rowNumber = sheet.getLastRow() + 1;
    sheet.getRange(rowNumber, 1, 1, lastCol).setValues([values]);
    formulasToSet.forEach(function (f) {
      sheet.getRange(rowNumber, f.column).setFormula(f.formula);
    });

    if (kind === 'gst') {
      if (savedInvoice) existingInvoices[savedInvoice] = true;
    } else if (kind === 'upcoming') {
      // Keep the old Sr No when it is still free; otherwise allocate the
      // next available number so onboarding after deletion cannot create a
      // duplicate when the customer is later restored.
      if (srColumn >= 0) {
        var restoredSr = parseInt(values[srColumn], 10);
        if (!(restoredSr > 0) || existingSrNumbers[String(restoredSr)]) {
          restoredSr = maxSrNumber + 1;
          values[srColumn] = restoredSr;
          sheet.getRange(rowNumber, srColumn + 1).setValue(restoredSr);
        }
        existingSrNumbers[String(restoredSr)] = true;
        maxSrNumber = Math.max(maxSrNumber, restoredSr);
      }
      if (savedName) existingNames[normName_(savedName)] = true;
      var expireIndex = binHeaderIndex_(currentHeaders, 'Expire Date');
      if (expireIndex >= 0) sheet.getRange(rowNumber, expireIndex + 1).setNumberFormat('dd/mm/yy');
    }
    appended++;
  });
  return appended;
}

function deleteSheetRows_(sheet, rowNumbers) {
  rowNumbers.slice().sort(function (a, b) { return b - a; }).forEach(function (row) {
    if (row >= 2 && row <= sheet.getLastRow()) sheet.deleteRow(row);
  });
}

function firstParentId_(driveItem) {
  var parents = driveItem.getParents();
  if (!parents.hasNext()) throw new Error('Cannot determine original Drive folder for "' + driveItem.getName() + '"');
  return parents.next().getId();
}

/**
 * Move a PDF archive folder while retaining its Drive ID and contents.
 * Folder.moveTo is the supported DriveApp operation for reparenting folders.
 */
function moveDriveFolder_(folderId, sourceParentId, destinationParent) {
  var folder = DriveApp.getFolderById(folderId);
  var source = DriveApp.getFolderById(sourceParentId);
  try {
    folder.moveTo(destinationParent);
  } catch (err) {
    // Best-effort rollback if the destination move fails partway through.
    try { folder.moveTo(source); } catch (rollbackErr) {}
    throw err;
  }
  return folder;
}

function moveBinDocsBack_(manifest, binDocsFolder) {
  var restored = [];
  (manifest.movedDocs || []).forEach(function (item) {
    var file = DriveApp.getFileById(item.id);
    var parent = DriveApp.getFolderById(item.originalParentId);
    file.moveTo(parent);
    restored.push(item);
  });
  return restored;
}

function restoreBinDocsToBin_(items, binDocsFolder) {
  (items || []).forEach(function (item) {
    try { DriveApp.getFileById(item.id).moveTo(binDocsFolder); } catch (e) {}
  });
}

function mapEntriesForIdentity_(map, identity, docIds) {
  var entries = {};
  Object.keys(map || {}).forEach(function (key) {
    var id = String(map[key] || '');
    if (identity.matches(key) || (id && docIds[id])) entries[key] = map[key];
  });
  return entries;
}

function removeMapEntries_(map, entries) {
  Object.keys(entries || {}).forEach(function (key) { delete map[key]; });
  return map;
}

function restoreMapEntries_(map, entries) {
  Object.keys(entries || {}).forEach(function (key) { map[key] = entries[key]; });
  return map;
}

function writeBinManifest_(folder, manifest) {
  var it = folder.getFilesByName('__restore_manifest.json');
  var file;
  if (it.hasNext()) file = it.next();
  else file = folder.createFile('__restore_manifest.json', '{}', MimeType.PLAIN_TEXT);
  file.setContent(JSON.stringify(manifest));
  return file.getId();
}

function readBinManifest_(folder) {
  var it = folder.getFilesByName('__restore_manifest.json');
  if (!it.hasNext()) throw new Error('Bin entry is missing its restore manifest');
  var file = it.next();
  var manifest;
  try { manifest = JSON.parse(file.getBlob().getDataAsString()); }
  catch (e) { throw new Error('Bin restore manifest is unreadable: ' + e.message); }
  if (!manifest || manifest.version !== 1) throw new Error('Unsupported or invalid Bin manifest');
  return { file: file, data: manifest };
}

function collectCustomerDocs_(identity, invoiceNumbers, docsMap, manualMap) {
  var ids = {};
  Object.keys(docsMap || {}).forEach(function (key) {
    if (identity.matches(key) && docsMap[key]) ids[String(docsMap[key])] = true;
  });
  Object.keys(manualMap || {}).forEach(function (key) {
    if (identity.matches(key) && manualMap[key]) ids[String(manualMap[key])] = true;
  });

  var byName = {};
  Object.keys(identity.keys).forEach(function (key) { byName[key] = true; });
  var invoiceLookup = {};
  (invoiceNumbers || []).forEach(function (number) {
    if (number) invoiceLookup[String(number)] = true;
  });

  var root = docsFolder_();
  var bin = customerBinFolder_(false);
  var pdfRoot = pdfsRoot_();
  var foundFiles = [];
  function walk(folder) {
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      if (file.getMimeType() !== MimeType.GOOGLE_DOCS) continue;
      var id = file.getId();
      var stem = String(file.getName() || '').replace(/\.(docx?)$/i, '').replace(/[_\s-]?BILL$/i, '');
      var match = !!byName[normName_(stem)];
      if (!match && Object.keys(invoiceLookup).length) {
        try {
          var text = docText_(id);
          Object.keys(invoiceLookup).some(function (invoiceNo) {
            if (text.indexOf(invoiceNo) >= 0) { match = true; return true; }
            return false;
          });
        } catch (e) { /* unmapped/unreadable docs are not selected by fuzzy name */ }
      }
      if (match) ids[id] = true;
    }

    var folders = folder.getFolders();
    while (folders.hasNext()) {
      var child = folders.next();
      if (child.getName().toUpperCase() === 'BIN') continue;
      if (pdfRoot && child.getId() === pdfRoot.getId()) continue;
      walk(child);
    }
  }
  walk(root);

  var out = [];
  Object.keys(ids).forEach(function (id) {
    var file = DriveApp.getFileById(id);
    if (file.getMimeType() !== MimeType.GOOGLE_DOCS)
      throw new Error('Mapped Drive item is not a Google Doc: ' + file.getName());
    out.push({ id: id, name: file.getName(), originalParentId: firstParentId_(file) });
  });
  return out;
}

/** Remove active GST/Upcoming rows only after Drive assets have been moved. */
function deleteCustomer_(body) {
  var customer = String(body.customer || '').trim();
  if (!customer) throw new Error('customer is required');

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another billing or customer operation is running. Try again.');
  var binEntry = null, movedDocs = [], movedPdfFolder = false, manifest = null;
  var ss, up, gst, upcomingRows = [], gstRows = [], docBin, pdfBin;
  try {
    ss = SpreadsheetApp.openById(SHEET_ID);
    up = ss.getSheetByName(UPCOMING_SHEET);
    gst = ss.getSheetByName(GST_SHEET);
    if (!up || !gst) throw new Error('Required spreadsheet tabs were not found');

    var identity = customerIdentityContext_(customer);
    var upVals = up.getDataRange().getValues();
    var gstVals = gst.getDataRange().getValues();
    var upNameIndex = colOf_(upVals[0], 'Customer name');
    var gstNameIndex = colOf_(gstVals[0], 'Customer Name');
    var invoiceIndex = colOf_(gstVals[0], 'Invoice No');
    var upRowNumbers = [], gstRowNumbers = [], invoiceNumbers = [];

    for (var u = 1; u < upVals.length; u++) {
      if (upNameIndex >= 0 && identity.matches(upVals[u][upNameIndex])) upRowNumbers.push(u + 1);
    }
    for (var g = 1; g < gstVals.length; g++) {
      if (gstNameIndex >= 0 && identity.matches(gstVals[g][gstNameIndex])) {
        gstRowNumbers.push(g + 1);
        if (invoiceIndex >= 0) invoiceNumbers.push(String(gstVals[g][invoiceIndex] || '').trim());
      }
    }
    var docsMap = loadJsonProperty_('DOCS_MAP');
    var manualMap = loadManualMap_();
    var hasMappedDoc = Object.keys(docsMap).some(function (key) {
      return identity.matches(key) && !!docsMap[key];
    }) || Object.keys(manualMap).some(function (key) {
      return identity.matches(key) && !!manualMap[key];
    });
    var pdfFolder = customerFolder_(customer, false);
    if (!upRowNumbers.length && !gstRowNumbers.length && !hasMappedDoc && !pdfFolder) {
      throw new Error('No active sheet records, mapped document or archived PDF folder was found for ' + customer);
    }

    var movedDocRecords = collectCustomerDocs_(identity, invoiceNumbers, docsMap, manualMap);
    var docIds = {};
    movedDocRecords.forEach(function (item) { docIds[item.id] = true; });
    var docMapEntries = mapEntriesForIdentity_(docsMap, identity, docIds);
    var manualMapEntries = mapEntriesForIdentity_(manualMap, identity, docIds);

    var pdfFolderRecord = null;
    if (pdfFolder) {
      pdfFolderRecord = {
        id: pdfFolder.getId(),
        name: pdfFolder.getName(),
        originalParentId: firstParentId_(pdfFolder)
      };
    }

    var gstSnapshot = snapshotSheetRows_(gst, gstRowNumbers);
    var upcomingSnapshot = snapshotSheetRows_(up, upRowNumbers);
    var binRoot = customerBinFolder_(true);
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    var entryName = sanitizeFile_(customer) + '__' + stamp + '__' + Utilities.getUuid().slice(0, 8);
    binEntry = binRoot.createFolder(entryName);
    docBin = binEntry.createFolder('Google Docs');
    pdfBin = binEntry.createFolder('PDFs');

    manifest = {
      version: 1,
      status: 'MOVING',
      deletedAt: new Date().toISOString(),
      customer: customer,
      identityAliases: Object.keys(identity.keys),
      invoiceCount: invoiceNumbers.length,
      invoiceNumbers: invoiceNumbers,
      upcomingSnapshot: upcomingSnapshot,
      gstSnapshot: gstSnapshot,
      movedDocs: [],
      pdfFolder: pdfFolderRecord,
      docMapEntries: docMapEntries,
      manualMapEntries: manualMapEntries,
      binEntryId: binEntry.getId()
    };
    writeBinManifest_(binEntry, manifest);

    movedDocRecords.forEach(function (item) {
      DriveApp.getFileById(item.id).moveTo(docBin);
      movedDocs.push(item);
      manifest.movedDocs.push(item);
      writeBinManifest_(binEntry, manifest);
    });

    if (pdfFolderRecord) {
      moveDriveFolder_(pdfFolderRecord.id, pdfFolderRecord.originalParentId, pdfBin);
      movedPdfFolder = true;
    }
    manifest.status = 'DRIVE_MOVED';
    writeBinManifest_(binEntry, manifest);

    deleteSheetRows_(up, upRowNumbers);
    deleteSheetRows_(gst, gstRowNumbers);

    docsMap = removeMapEntries_(docsMap, docMapEntries);
    manualMap = removeMapEntries_(manualMap, manualMapEntries);
    saveDocsMap_(docsMap);
    saveManualMap_(manualMap);

    manifest.status = 'DELETED';
    manifest.deletedAt = new Date().toISOString();
    writeBinManifest_(binEntry, manifest);
    SpreadsheetApp.flush();

    return {
      ok: true,
      customer: customer,
      deleted_id: binEntry.getId(),
      moved_documents: movedDocs.length,
      moved_pdf_folder: !!pdfFolderRecord,
      removed_upcoming_rows: upRowNumbers.length,
      removed_gst_rows: gstRowNumbers.length
    };
  } catch (err) {
    // Best-effort rollback: put sheet rows, docs, PDF folder and mappings back.
    try {
      if (up && manifest) restoreSheetSnapshot_(up, manifest.upcomingSnapshot, 'upcoming');
      if (gst && manifest) restoreSheetSnapshot_(gst, manifest.gstSnapshot, 'gst');
      if (docBin) restoreBinDocsToBin_(movedDocs, docBin); // re-home first, then original parents below
      movedDocs.slice().reverse().forEach(function (item) {
        try { DriveApp.getFileById(item.id).moveTo(DriveApp.getFolderById(item.originalParentId)); } catch (e1) {}
      });
      if (movedPdfFolder && manifest && manifest.pdfFolder) {
        try {
          moveDriveFolder_(manifest.pdfFolder.id, pdfBin.getId(),
            DriveApp.getFolderById(manifest.pdfFolder.originalParentId));
        } catch (e2) {}
      }
      if (manifest) {
        var dm = loadJsonProperty_('DOCS_MAP');
        var mm = loadManualMap_();
        saveDocsMap_(restoreMapEntries_(dm, manifest.docMapEntries));
        saveManualMap_(restoreMapEntries_(mm, manifest.manualMapEntries));
      }
      if (binEntry) binEntry.setTrashed(true);
    } catch (rollbackErr) {
      throw new Error(String(err && err.message || err) +
        ' | Rollback also reported: ' + String(rollbackErr && rollbackErr.message || rollbackErr) +
        '. Check the BIN folder before retrying.');
    }
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function listDeletedCustomers_(body) {
  var binRoot = customerBinFolder_(false);
  if (!binRoot) return { ok: true, items: [] };
  var items = [], folders = binRoot.getFolders();
  while (folders.hasNext()) {
    var folder = folders.next();
    try {
      var found = readBinManifest_(folder);
      var manifest = found.data;
      if (manifest.status !== 'DELETED') continue;
      var lastExpire = '';
      if (manifest.upcomingSnapshot && manifest.upcomingSnapshot.rows.length) {
        var headers = manifest.upcomingSnapshot.headers;
        var expIndex = binHeaderIndex_(headers, 'Expire Date');
        if (expIndex >= 0) {
          var expiry = unpackBinCell_(manifest.upcomingSnapshot.rows[0].values[expIndex]);
          if (expiry instanceof Date) lastExpire = fmtDMY_(expiry);
          else lastExpire = String(expiry || '');
        }
      }
      items.push({
        id: folder.getId(),
        customer: manifest.customer,
        deleted_at: manifest.deletedAt,
        invoice_count: (manifest.invoiceNumbers || []).length,
        upcoming_rows: manifest.upcomingSnapshot ? manifest.upcomingSnapshot.rows.length : 0,
        expire_date: lastExpire
      });
    } catch (e) {
      // Ignore unrelated/manual folders in BIN; only valid manifests are app entries.
    }
  }
  items.sort(function (a, b) { return String(b.deleted_at).localeCompare(String(a.deleted_at)); });
  return { ok: true, items: items };
}

function binEntryById_(entryId) {
  var id = String(entryId || '').trim();
  if (!id) throw new Error('deleted customer id is required');
  var root = customerBinFolder_(false);
  if (!root) throw new Error('BIN folder does not exist');
  var folder;
  try { folder = DriveApp.getFolderById(id); }
  catch (e) { throw new Error('Deleted customer folder was not found in BIN'); }
  var parents = folder.getParents();
  var belongs = false;
  while (parents.hasNext()) if (parents.next().getId() === root.getId()) belongs = true;
  if (!belongs) throw new Error('The selected folder is not a direct entry in BIN');
  var found = readBinManifest_(folder);
  if (found.data.status !== 'DELETED') throw new Error('This Bin entry is not in a restorable deleted state');
  return { folder: folder, manifestFile: found.file, manifest: found.data };
}

function validateRestoreHasNoConflict_(manifest) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var up = ss.getSheetByName(UPCOMING_SHEET);
  var gst = ss.getSheetByName(GST_SHEET);
  var upVals = up.getDataRange().getValues(), gstVals = gst.getDataRange().getValues();
  var identity = { matches: function (value) {
    var key = normName_(value);
    if (!key) return false;
    if (key === normName_(manifest.customer)) return true;
    return (manifest.identityAliases || []).indexOf(key) >= 0 ||
      canonicalCustomerName_(value) === canonicalCustomerName_(manifest.customer);
  }};
  var upNameIndex = colOf_(upVals[0], 'Customer name');
  for (var r = 1; r < upVals.length; r++) {
    if (upNameIndex >= 0 && identity.matches(upVals[r][upNameIndex]))
      throw new Error('A customer with this name is already active. Rename or remove the active record before restoring.');
  }
  var gstNameIndex = colOf_(gstVals[0], 'Customer Name');
  for (var gr = 1; gr < gstVals.length; gr++) {
    if (gstNameIndex >= 0 && identity.matches(gstVals[gr][gstNameIndex]))
      throw new Error('A GST RECORDS entry for this customer is already active. Remove or rename the active records before restoring.');
  }

  var invIndex = colOf_(gstVals[0], 'Invoice No');
  var existing = {};
  for (var g = 1; g < gstVals.length; g++) {
    if (invIndex >= 0) existing[String(gstVals[g][invIndex] || '').trim()] = true;
  }
  (manifest.invoiceNumbers || []).forEach(function (invoiceNo) {
    if (existing[invoiceNo]) throw new Error('Invoice ' + invoiceNo + ' already exists in active GST RECORDS. Restore was stopped to avoid duplicates.');
  });

  var docsMap = loadJsonProperty_('DOCS_MAP');
  var manualMap = loadManualMap_();
  var hasActiveDocument = Object.keys(docsMap).some(function (key) {
    return identity.matches(key) && !!docsMap[key];
  }) || Object.keys(manualMap).some(function (key) {
    return identity.matches(key) && !!manualMap[key];
  });
  if (hasActiveDocument)
    throw new Error('A Google Doc mapping for this customer is already active. Resolve the active mapping before restoring.');
}

function restoreDeletedCustomer_(body) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another billing or customer operation is running. Try again.');
  var binInfo = null, movedBackDocs = [], movedBackPdf = false, rowsRestored = false, mapsRestored = false;
  var binDocs = null, binPdfs = null;
  try {
    binInfo = binEntryById_(body.deleted_id);
    var manifest = binInfo.manifest;
    validateRestoreHasNoConflict_(manifest);

    var sub = binInfo.folder.getFolders();
    while (sub.hasNext()) {
      var child = sub.next();
      if (child.getName() === 'Google Docs') binDocs = child;
      if (child.getName() === 'PDFs') binPdfs = child;
    }

    // Return files to their original Drive parents before exposing sheet rows.
    (manifest.movedDocs || []).forEach(function (item) {
      var file = DriveApp.getFileById(item.id);
      file.moveTo(DriveApp.getFolderById(item.originalParentId));
      movedBackDocs.push(item);
    });
    if (manifest.pdfFolder) {
      moveDriveFolder_(manifest.pdfFolder.id, binPdfs.getId(),
        DriveApp.getFolderById(manifest.pdfFolder.originalParentId));
      movedBackPdf = true;
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var up = ss.getSheetByName(UPCOMING_SHEET);
    var gst = ss.getSheetByName(GST_SHEET);
    // Mark rollback as necessary before restoring the first row: a failure
    // midway through either sheet must still remove any partially restored rows.
    rowsRestored = true;
    restoreSheetSnapshot_(gst, manifest.gstSnapshot, 'gst');
    restoreSheetSnapshot_(up, manifest.upcomingSnapshot, 'upcoming');

    var docsMap = loadJsonProperty_('DOCS_MAP');
    var manualMap = loadManualMap_();
    saveDocsMap_(restoreMapEntries_(docsMap, manifest.docMapEntries));
    saveManualMap_(restoreMapEntries_(manualMap, manifest.manualMapEntries));
    mapsRestored = true;

    var headers = up.getRange(1, 1, 1, up.getLastColumn()).getValues()[0];
    var expireIndex = colOf_(headers, 'Expire Date');
    if (expireIndex >= 0) sortUpcomingByExpire_(up, expireIndex);
    SpreadsheetApp.flush();

    var restoredName = manifest.customer;
    var result = {
      ok: true, customer: restoredName,
      restored_documents: (manifest.movedDocs || []).length,
      restored_pdf_folder: !!manifest.pdfFolder,
      restored_upcoming_rows: manifest.upcomingSnapshot ? manifest.upcomingSnapshot.rows.length : 0,
      restored_gst_rows: manifest.gstSnapshot ? manifest.gstSnapshot.rows.length : 0
    };

    // Mark before removing the Bin wrapper. If removing the wrapper fails, it
    // must not reappear as an active restore option.
    manifest.status = 'RESTORED';
    writeBinManifest_(binInfo.folder, manifest);
    try { binInfo.folder.setTrashed(true); } catch (cleanupErr) {}
    return result;
  } catch (err) {
    // Keep the customer safely in the Bin if restore did not complete.
    if (movedBackDocs.length && binDocs) {
      movedBackDocs.forEach(function (item) {
        try { DriveApp.getFileById(item.id).moveTo(binDocs); } catch (e1) {}
      });
    }
    if (movedBackPdf && binPdfs && binInfo && binInfo.manifest.pdfFolder) {
      try {
        var currentParentId = firstParentId_(DriveApp.getFolderById(binInfo.manifest.pdfFolder.id));
        moveDriveFolder_(binInfo.manifest.pdfFolder.id, currentParentId, binPdfs);
      } catch (e2) {}
    }
    if (rowsRestored && binInfo) {
      try {
        var ss2 = SpreadsheetApp.openById(SHEET_ID);
        var up2 = ss2.getSheetByName(UPCOMING_SHEET);
        var gst2 = ss2.getSheetByName(GST_SHEET);
        var identity = { matches: function (v) {
          return normName_(v) === normName_(binInfo.manifest.customer) ||
            (binInfo.manifest.identityAliases || []).indexOf(normName_(v)) >= 0;
        }};
        var upVals2 = up2.getDataRange().getValues(), gstVals2 = gst2.getDataRange().getValues();
        var upName2 = colOf_(upVals2[0], 'Customer name'), gstInv2 = colOf_(gstVals2[0], 'Invoice No');
        var upRows = [], gstRows = [], invs = {};
        (binInfo.manifest.invoiceNumbers || []).forEach(function (v) { invs[v] = true; });
        for (var ur = 1; ur < upVals2.length; ur++) if (upName2 >= 0 && identity.matches(upVals2[ur][upName2])) upRows.push(ur + 1);
        for (var gr = 1; gr < gstVals2.length; gr++) if (gstInv2 >= 0 && invs[String(gstVals2[gr][gstInv2] || '').trim()]) gstRows.push(gr + 1);
        deleteSheetRows_(up2, upRows);
        deleteSheetRows_(gst2, gstRows);
      } catch (e3) {}
    }
    if (mapsRestored && binInfo) {
      try {
        var currentDocsMap = loadJsonProperty_('DOCS_MAP');
        var currentManualMap = loadManualMap_();
        saveDocsMap_(removeMapEntries_(currentDocsMap, binInfo.manifest.docMapEntries));
        saveManualMap_(removeMapEntries_(currentManualMap, binInfo.manifest.manualMapEntries));
      } catch (e4) {}
    }
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/** Permanently remove a Drive item through Drive API v3 (not merely trash it). */
function permanentlyDeleteDriveItem_(fileId) {
  var url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) + '?supportsAllDrives=true';
  var response = UrlFetchApp.fetch(url, {
    method: 'delete',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  if (code !== 200 && code !== 204 && code !== 404)
    throw new Error('Drive permanent delete failed for ' + fileId + ' (HTTP ' + code + '): ' +
      String(response.getContentText() || '').slice(0, 250));
}

function permanentlyDeleteDriveTree_(folder) {
  var files = folder.getFiles();
  while (files.hasNext()) permanentlyDeleteDriveItem_(files.next().getId());
  var folders = folder.getFolders();
  while (folders.hasNext()) permanentlyDeleteDriveTree_(folders.next());
  permanentlyDeleteDriveItem_(folder.getId());
}

function purgeDeletedCustomer_(body) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another billing or customer operation is running. Try again.');
  try {
    var info = binEntryById_(body.deleted_id);
    var customer = info.manifest.customer;
    var aliasesToRemove = {};
    (info.manifest.identityAliases || []).forEach(function (key) { aliasesToRemove[String(key)] = true; });
    aliasesToRemove[normName_(customer)] = true;

    // First remove the Drive tree. If the API reports a failure, keep the
    // manifest and all alias metadata intact so the customer can be restored.
    permanentlyDeleteDriveTree_(info.folder);

    var folderAliases = loadCustomerFolderAliases_();
    Object.keys(folderAliases).forEach(function (key) {
      if (aliasesToRemove[key]) delete folderAliases[key];
    });
    PROPS.setProperty('CUSTOMER_FOLDER_ALIASES', JSON.stringify(folderAliases));

    var nameAliases = loadJsonProperty_('CUSTOMER_NAME_ALIASES');
    Object.keys(nameAliases).forEach(function (key) {
      if (aliasesToRemove[key] || aliasesToRemove[normName_(nameAliases[key])]) delete nameAliases[key];
    });
    PROPS.setProperty('CUSTOMER_NAME_ALIASES', JSON.stringify(nameAliases));
    customerNameAliasCache_ = nameAliases;

    return { ok: true, permanently_deleted: true, customer: customer };
  } finally {
    lock.releaseLock();
  }
}

function syncDocs_(body) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var up = ss.getSheetByName(UPCOMING_SHEET);
  var upVals = up.getDataRange().getValues();
  var iName = colOf_(upVals[0], 'Customer name');
  var customers = [];
  for (var r = 1; r < upVals.length; r++) {
    var n = String(upVals[r][iName] || '').trim();
    if (n) customers.push(n);
  }
  var files = [];
  function walk(folder, path) {
    var fit = folder.getFiles();
    while (fit.hasNext()) {
      var f = fit.next();
      files.push({ name: f.getName(), id: f.getId(), path: path,
                   isGoogleDoc: f.getMimeType() === MimeType.GOOGLE_DOCS });
    }
    var dit = folder.getFolders();
    while (dit.hasNext()) {
      var d = dit.next();
      // The recycle bin is nested under Gst_Bills but must never be scanned
      // by sync_docs; deleted customer Docs must stay out of the active map.
      if (d.getName().toUpperCase() === 'BIN') continue;
      walk(d, path + d.getName() + '/');
    }
  }
  walk(docsFolder_(), '');
  var map = {}, report = [];
  var manual = loadManualMap_();
  files.forEach(function (f) {
    var best = null, bestScore = 0;
    customers.forEach(function (c) {
      var s = nameScore_(f.name.replace(/\.docx?$/i, ''), c);
      if (s > bestScore) { bestScore = s; best = c; }
    });
    var entry = { doc: f.name, id: f.id, path: f.path, isGoogleDoc: f.isGoogleDoc,
                  match: bestScore >= 0.8 ? best : '(no clear match)',
                  score: Math.round(bestScore * 100) / 100 };
    if (bestScore >= 0.8 && best) {
      if (manual[best]) {
        // customer is bound manually — never auto-override it
        entry.status = 'manual mapping kept: ' + manual[best];
      } else if (f.isGoogleDoc && !map[best]) {
        entry.status = 'mapped';
        map[best] = f.id;
      } else if (f.isGoogleDoc && map[best]) {
        entry.status = 'duplicate match (first kept)';
      } else {
        entry.status = 'needs conversion (open as Google Doc)';
      }
    } else {
      entry.status = 'unmatched';
    }
    report.push(entry);
  });
  for (var mc in manual) map[mc] = manual[mc];   // manual wins, always
  saveDocsMap_(map);
  return { ok: true, docs_found: files.length, customers: customers.length,
           mapped: Object.keys(map).length, map: map, report: report };
}

/** map_doc: manually bind one Google Doc to one customer (e.g. files whose
 *  names the auto-matcher can't guess: NARENDRAA_BILL, STNHQ_BILL,
 *  RANKA_BUNGLOW_BILL). The binding is stored in DOCS_MAP_MANUAL and survives
 *  every sync_docs run — manual always wins over auto-match. */
function mapDoc_(body) {
  var cust = String(body.customer || '').trim();
  var docId = String(body.doc_id || '').trim();
  if (!cust || !docId) throw new Error('customer and doc_id are required');
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var up = ss.getSheetByName(UPCOMING_SHEET);
  var vals = up.getDataRange().getValues();
  var iName = colOf_(vals[0], 'Customer name');
  var found = '';
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][iName] || '').trim().toLowerCase() === cust.toLowerCase()) {
      found = String(vals[r][iName] || '').trim();
      break;
    }
  }
  if (!found) {
    // fall back: allow names that exist only in GST RECORDS (e.g. service
    // sub-records like "… Switch" / "… Router" with no UPCOMING BILLS row)
    var gst2 = ss.getSheetByName(GST_SHEET);
    var gv2 = gst2.getDataRange().getValues();
    var iCust2 = colOf_(gv2[0], 'Customer Name');
    if (iCust2 >= 0) {
      for (var g2 = 1; g2 < gv2.length; g2++) {
        if (String(gv2[g2][iCust2] || '').trim().toLowerCase() === cust.toLowerCase()) {
          found = String(gv2[g2][iCust2] || '').trim();
          break;
        }
      }
    }
  }
  if (!found) throw new Error('customer not found in UPCOMING BILLS or GST RECORDS: ' + cust);
  var f = DriveApp.getFileById(docId);
  if (f.getMimeType() !== MimeType.GOOGLE_DOCS) {
    throw new Error('not a Google Doc yet — File > Save as Google Docs first (mime: ' +
      f.getMimeType() + ')');
  }
  var manual = loadManualMap_();
  manual[found] = docId;
  saveManualMap_(manual);
  var merged = {};
  try { merged = JSON.parse(PROPS.getProperty('DOCS_MAP') || '{}'); } catch (e) { merged = {}; }
  merged[found] = docId;
  saveDocsMap_(merged);
  var starts = '';
  try { starts = DocumentApp.openById(docId).getBody().getText().slice(0, 100); } catch (e2) {}
  return { ok: true, customer: found, doc_id: docId, doc_name: f.getName(),
           mapped: true, doc_starts: starts };
}

/** Doc id -> previous-bill value pairs for the in-place edit. Old values are
 *  reconstructed the same way makePdf_ does: last GST record + the Duration
 *  column (which still holds the last bill's duration) + the customer's
 *  current Expire Date (= that bill's period end). */
function docEditPairs_(last, oldExp, lastMonths, bill) {
  var oldStart = (oldExp && lastMonths > 0) ? addDays_(addMonths_(oldExp, -lastMonths), 1) : null;
  var oldPeriod = oldStart ? (fmtDMY_(oldStart) + ' to ' + fmtDMY_(oldExp)) : '';
  var oldDate = String(last.invoiceDate || '').replace(/\//g, '-');   // docs print dd-mm-yy
  var newDate = String(bill.invoiceDate || '').replace(/\//g, '-');
  function words(n) { return rupeeWords_(n) + ' RUPEES ONLY/_'; }
  return [
    { name: 'Invoice No',   old: last.invoiceNo,  nw: bill.invoiceNo },
    { name: 'Invoice Date', old: oldDate,         nw: newDate },
    { name: 'Bill Period',  old: oldPeriod,       nw: bill.periodStart + ' to ' + bill.periodEnd },
    { name: 'Amount',       old: money_(last.amount), nw: money_(bill.amount) },
    { name: 'GST cells',    old: money_(last.cgst),   nw: money_(bill.cgst) },
    { name: 'Total',        old: money_(last.total),  nw: money_(bill.total) },
    { name: 'Words',        old: words(last.total),   nw: words(bill.total) },
    { name: 'Description',  old: String(last.desc || ''), nw: String(bill.desc || ''), warn: true }
  ];
}

/* ------------------------------------------------------- doc editing ------ */

function escRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Read printed invoice date with slash or hyphen separators. */
function invoiceDateText_(text) {
  var m = String(text || '').match(/Invoice\s*Date\s*[:\-]*\s*(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/i);
  return m ? m[1] : '';
}

/**
 * Read the printed invoice date from the actual Google Docs structure.
 * In this invoice layout, "Invoice Date" and its value occupy adjacent table
 * cells. getBody().getText() flattens tables and inserts separators between
 * cells, so matching the label and value in one text string is unreliable.
 *
 * Scan each table row, find the cell containing the label, then read the date
 * from that cell or a cell to its right in the SAME row. Nested tables,
 * document body, header and footer are supported.
 */
function invoiceDateTextInDoc_(docId) {
  var doc = DocumentApp.openById(docId);
  var roots = [doc.getBody(), doc.getHeader(), doc.getFooter()];

  function dateIn_(value) {
    var m = String(value || '').match(/\b\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}\b/);
    return m ? m[0] : '';
  }

  function scan_(element) {
    if (!element || typeof element.getType !== 'function') return '';
    var type = element.getType();

    if (type === DocumentApp.ElementType.TABLE) {
      var table = element;
      for (var r = 0; r < table.getNumRows(); r++) {
        var row = table.getRow(r);
        var cells = [];
        for (var c = 0; c < row.getNumCells(); c++) cells.push(row.getCell(c));

        for (var c0 = 0; c0 < cells.length; c0++) {
          var cellText = cells[c0].getText();
          if (!/Invoice\s*Date/i.test(cellText)) continue;

          var sameCellDate = dateIn_(cellText);
          if (sameCellDate) return sameCellDate;

          // The value normally sits immediately to the right. Continue
          // within this row only; never borrow a date from another row.
          for (var c1 = c0 + 1; c1 < cells.length; c1++) {
            var adjacentDate = dateIn_(cells[c1].getText());
            if (adjacentDate) return adjacentDate;
          }
        }

        // Metadata tables may be nested inside an outer invoice table.
        for (var c2 = 0; c2 < cells.length; c2++) {
          var nested = scan_(cells[c2]);
          if (nested) return nested;
        }
      }
      return '';
    }

    if (typeof element.getNumChildren === 'function') {
      for (var i = 0; i < element.getNumChildren(); i++) {
        var found = scan_(element.getChild(i));
        if (found) return found;
      }
    }
    return '';
  }

  for (var i0 = 0; i0 < roots.length; i0++) {
    if (!roots[i0]) continue;
    var foundDate = scan_(roots[i0]);
    if (foundDate) return foundDate;
  }

  // Fallback for documents where label and value share one text block.
  return invoiceDateText_(docText_(docId));
}

/**
 * Extract the actual printed buyer name from the "Buyer (Bill to)" cell.
 * This is important after a Doc/PDF-only rename: GST RECORDS intentionally
 * retains its customer key unless the user opts into a sheet rename.
 */
function buyerNameInDoc_(docId) {
  var doc = DocumentApp.openById(docId);
  var roots = [doc.getBody(), doc.getHeader(), doc.getFooter()];

  function cleanCandidate_(value) {
    var v = String(value || '').replace(/\u00A0/g, ' ').trim();
    if (!v || /^(GSTIN|GSTIN \/ UIN|State Name|SAC Code|Invoice Date|Invoice No)\b/i.test(v)) return '';
    return v;
  }

  function fromCellText_(value) {
    var lines = String(value || '').replace(/\u00A0/g, ' ').split(/\r?\n/);
    for (var i = 0; i < lines.length; i++) {
      if (!/Buyer\s*\(\s*Bill\s*to\s*\)/i.test(lines[i])) continue;
      var sameLine = cleanCandidate_(lines[i].replace(/.*Buyer\s*\(\s*Bill\s*to\s*\)\s*/i, ''));
      if (sameLine) return sameLine;
      for (var j = i + 1; j < lines.length; j++) {
        var next = cleanCandidate_(lines[j]);
        if (next) return next;
      }
    }
    return '';
  }

  function scan_(element) {
    if (!element || typeof element.getType !== 'function') return '';
    if (element.getType() === DocumentApp.ElementType.TABLE) {
      var table = element;
      for (var r = 0; r < table.getNumRows(); r++) {
        var row = table.getRow(r), cells = [];
        for (var c = 0; c < row.getNumCells(); c++) cells.push(row.getCell(c));
        for (var c0 = 0; c0 < cells.length; c0++) {
          var found = fromCellText_(cells[c0].getText());
          if (found) return found;
          if (/Buyer\s*\(\s*Bill\s*to\s*\)/i.test(cells[c0].getText())) {
            for (var c1 = c0 + 1; c1 < cells.length; c1++) {
              var adjacent = cleanCandidate_(cells[c1].getText());
              if (adjacent) return adjacent.split(/\r?\n/)[0].trim();
            }
          }
        }
        for (var c2 = 0; c2 < cells.length; c2++) {
          var nested = scan_(cells[c2]);
          if (nested) return nested;
        }
      }
      return '';
    }
    if (typeof element.getNumChildren === 'function') {
      for (var i0 = 0; i0 < element.getNumChildren(); i0++) {
        var child = scan_(element.getChild(i0));
        if (child) return child;
      }
    }
    return '';
  }

  for (var i1 = 0; i1 < roots.length; i1++) {
    if (!roots[i1]) continue;
    var result = scan_(roots[i1]);
    if (result) return result;
  }
  var flat = docText_(docId);
  var m = flat.match(/Buyer\s*\(\s*Bill\s*to\s*\)\s*[\r\n]+([^\r\n]+)/i);
  return m ? cleanCandidate_(m[1]) : '';
}

/** Preserve the separator used by the existing customer document. */
function formatInvoiceDateLike_(date, sample) {
  var out = date instanceof Date ? fmtShort_(date) : String(date || '').trim();
  out = out.replace(/-/g, '/');
  return String(sample || '').indexOf('-') >= 0 ? out.replace(/\//g, '-') : out;
}

/* ---------------------------------------------------------------------------
 * Editing engine: pure Apps Script (DocumentApp). No Docs API, no REST calls
 * for editing.
 *
 * HISTORY: uploaded .docx files hard-abort on ANY server-side write
 * ("The document is inaccessible") — that was the failure we chased. The fix
 * was to make every bill doc a TRUE Google Doc: the template was converted
 * once (docx -> native doc) and every customer doc is a native copy of it.
 * On native Google Docs the DocumentApp path is the proven, working engine,
 * so all editing uses DocumentApp alone.
 *
 * RULE: if a doc ever fails to edit with "The document is inaccessible",
 * it is an Office-file state — re-create it from the Google Doc template
 * (rebuild_doc), never edit the .docx blob directly.
 * ------------------------------------------------------------------------- */

/** Full text of a doc for verification — DocumentApp read. */
function docText_(docId) {
  var doc = DocumentApp.openById(docId);
  var parts = [doc.getBody().getText()];
  var header = doc.getHeader(), footer = doc.getFooter();
  if (header) parts.push(header.getText());
  if (footer) parts.push(footer.getText());
  return parts.join('\n');
}

/** One text replacement. DocumentApp.replaceText treats the pattern as regex
 *  and "$" in the replacement as a capture reference — escape both sides. */
function replaceTextInDoc_(docId, oldText, newText) {
  var doc = DocumentApp.openById(docId);
  try {
    var targets = [doc.getBody(), doc.getHeader(), doc.getFooter()];
    for (var i = 0; i < targets.length; i++) {
      if (targets[i]) targets[i].replaceText(
        escRegex_(oldText), String(newText).replace(/\$/g, '$$'));
    }
  } finally {
    doc.saveAndClose();
  }
}

/** Best-effort description replacement: whole string first, then line-wise. */
function replaceDesc_(docId, oldDesc, newDesc) {
  if (!oldDesc || !newDesc || oldDesc === newDesc) return 'unchanged';
  var text = docText_(docId);
  if (text.indexOf(oldDesc) >= 0) {
    replaceTextInDoc_(docId, oldDesc, newDesc);
    return 'replaced';
  }
  var oldLines = oldDesc.split('\n').map(function (s) { return s.trim(); }).filter(Boolean);
  var newLines = newDesc.split('\n').map(function (s) { return s.trim(); }).filter(Boolean);
  if (!oldLines.length) return 'NOT FOUND — left unchanged';
  if (oldLines.length === 1 && newLines.length > 1) {
    // one paragraph in the doc, two-line description now — merge onto one line
    if (text.indexOf(oldLines[0]) < 0) return 'NOT FOUND — left unchanged';
    replaceTextInDoc_(docId, oldLines[0], newLines.join(' '));
    return 'replaced (merged to one line)';
  }
  var max = Math.max(oldLines.length, newLines.length);
  var misses = [];
  for (var i = 0; i < max; i++) {
    var o = i < oldLines.length ? oldLines[i] : null;
    var w = i < newLines.length ? newLines[i] : '';
    if (o === null) continue;                       // extra new lines have no anchor
    if (o === w) continue;
    if (text.indexOf(o) < 0) { misses.push(o); continue; }
    replaceTextInDoc_(docId, o, w);
  }
  return misses.length ? 'partially replaced, not found: ' + misses.join(' | ')
                       : 'replaced (line-wise)';
}

/** Applies old->new pairs to the customer doc IN PLACE, with verification.
 *  Edits go through replaceTextInDoc_ (Docs API in rest mode).
 *  - a pair whose "old" is absent but whose "new" is already present counts as
 *    already updated (safe retry after a partial previous attempt)
 *  - a critical pair missing entirely -> hard error, nothing changed
 *  - after applying, every new value must be present in the doc or the edit
 *    is reverted and an error thrown
 *  Returns {edited: [..], skipped: [..]}.
 */
function editCustomerDoc_(docId, pairs) {
  var CRITICAL = ['Invoice No', 'Invoice Date', 'Bill Period',
                  'Amount', 'GST cells', 'Total', 'Words'];
  var before = docText_(docId);
  var applied = [], skipped = [], missing = [];

  pairs.forEach(function (p) {
    if (p.name === 'Description') return;           // handled separately below
    if (!p.old || p.old === p.nw) { skipped.push(p.name + ' (unchanged)'); return; }
    if (before.indexOf(p.old) >= 0) applied.push(p);
    else if (before.indexOf(p.nw) >= 0) skipped.push(p.name + ' (already updated)');
    else missing.push(p);
  });

  var critMissing = missing.filter(function (p) { return CRITICAL.indexOf(p.name) >= 0; });
  if (critMissing.length) {
    var allAlreadyNew = critMissing.every(function (p) { return before.indexOf(p.nw) >= 0; });
    if (!allAlreadyNew) {
      throw new Error('customer doc does not contain the previous bill — fields not found: ' +
        critMissing.map(function (p) { return p.name + ' ("' + String(p.old).slice(0, 40) + '")'; }).join('; ') +
        '. Open the doc and check it matches the last GST RECORDS entry, or rebuild it (rebuild_doc).');
    }
    // every critical pair is already updated in the doc — safe to continue
    critMissing.forEach(function (p) { skipped.push(p.name + ' (already updated)'); });
    missing = missing.filter(function (p) { return CRITICAL.indexOf(p.name) < 0; });
  }
  missing.forEach(function (p) { skipped.push(p.name + ' (not found — left unchanged)'); });

  // longer values first so a shorter old value can never sit inside a longer one
  applied.sort(function (a, b) { return String(b.old).length - String(a.old).length; });
  applied.forEach(function (p) { replaceTextInDoc_(docId, p.old, p.nw); });

  var descPair = pairs.filter(function (p) { return p.name === 'Description'; })[0];
  var descResult = '';
  if (descPair) descResult = replaceDesc_(docId, descPair.old, descPair.nw);

  var after = docText_(docId);
  var failures = applied.filter(function (p) { return after.indexOf(p.nw) < 0; });
  if (failures.length) {
    failures.forEach(function (p) {                 // best-effort revert
      try { replaceTextInDoc_(docId, p.nw, p.old); } catch (e) { /* keep going */ }
    });
    throw new Error('doc edit verification failed for: ' +
      failures.map(function (p) { return p.name; }).join('; ') + ' — edit reverted');
  }
  return { edited: applied.map(function (p) { return p.name; }),
           skipped: skipped, desc: descResult };
}

/* ------------------------------------------------- export + archive ------- */

/** doc -> PDF blob, by doc ID. Web export endpoint first (independent
 *  pipeline — survives hard failures of the Docs text session); DocumentApp
 *  export as fallback (reads/exports are proven safe on problem docs). */
function exportPdf_(docId) {
  var firstErr = '';

  // PRIMARY: convert the current Google Doc directly. The Docs /export URL
  // can return a cached pre-edit render even when the document text is new.
  try {
    // Google Docs Document objects do not expose Drive File conversion.
    // Convert the saved Drive file itself to PDF.
    var blob = DriveApp.getFileById(docId).getAs(MimeType.PDF);
    if (blob && blob.getContentType() === 'application/pdf' &&
        blob.getBytes().length > 1000) return blob;
    firstErr = 'DocumentApp export: ' +
      (blob ? ('invalid/tiny PDF: ' + blob.getBytes().length + ' bytes') : 'null blob');
  } catch (e) {
    firstErr = 'DocumentApp export: ' + String(e && e.message || e).slice(0, 110);
  }

  // FALLBACK only: use the export URL if direct conversion fails.
  try {
    var resp = UrlFetchApp.fetch('https://docs.google.com/document/d/' + docId +
      '/export?format=pdf', {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true, followRedirects: true
    });
    var out = resp.getBlob();
    if (resp.getResponseCode() === 200 && out.getContentType() === 'application/pdf' &&
        out.getBytes().length > 1000) return out;
    firstErr += ' — export endpoint: HTTP ' + resp.getResponseCode() + ' ' + out.getContentType();
  } catch (e) {
    firstErr += ' — export endpoint: ' + String(e && e.message || e).slice(0, 110);
  }
  throw new Error('PDF export failed — ' + firstErr);
}

/** Exports the doc to PDF and archives it under BILLS/<Customer>/.
 *  Root cause found 17/11/26: Google's export endpoint serves a CACHED,
 *  PRE-EDIT render for MINUTES after a doc edit (twice confirmed — archived
 *  PDFs were byte-identical to the stale render while the doc held new
 *  values). No settle delay is long enough, and bills are shared on WhatsApp
 *  immediately. So the PRIMARY path is the copy trick: make a throwaway COPY
 *  of the doc and export that — a fresh copy has no render cache, so its
 *  first export is always the post-edit content. (Trying the main doc's
 *  export first was pointless: right after an edit its cache is essentially
 *  always stale, so the cheap call failed its freshness check every time and
 *  only added delay.) The pre-edit fingerprint verification and background
 *  trigger self-heal stay as safety nets. */
function exportAndArchive_(docId, custName, invoiceNo, dateStr, baselineHash) {
  var fname = sanitizeFile_(invoiceNo) + '_' + sanitizeFile_(dateStr) + '.pdf';
  var folder = customerFolder_(custName, true);
  var blob = null, settled = true;
  try { blob = exportFreshViaCopy_(docId); }          // primary: cache-immune
  catch (e) { blob = null; }
  if (!blob) blob = exportPdf_(docId);                // rare fallback
  // verify against the pre-edit fingerprint; if somehow still identical,
  // retry a few times (direct, then copy again) before archiving
  if (baselineHash && blobHash_(blob) === baselineHash) {
    settled = false;
    for (var attempt = 0; attempt < 3 && !settled; attempt++) {
      Utilities.sleep(4000);
      blob = exportFreshViaCopy_(docId);
      if (blobHash_(blob) !== baselineHash) { settled = true; break; }
      try {
        var viaCopy = exportFreshViaCopy_(docId);
        if (blobHash_(viaCopy) !== baselineHash) { blob = viaCopy; settled = true; }
      } catch (e) { /* next round */ }
    }
  }
  blob.setName(fname);
  var file = folder.createFile(blob);
  if (!settled) {
    try {
      enqueuePdfRefresh_({ fileId: file.getId(), docId: docId,
                           fname: fname, baseline: baselineHash, attempts: 0 });
    } catch (queueError) {
      try { file.setTrashed(true); } catch (cleanupError) {}
      throw queueError;
    }
  }
  return { blob: blob, file: fname, fileId: file.getId(), folderUrl: folder.getUrl(), pdf_settled: settled };
}

/** Export a THROWAWAY COPY of the doc: a brand-new doc has no cached render,
 *  so its first export reflects the current (post-edit) content. The copy is
 *  trashed immediately after. */
function exportFreshViaCopy_(docId) {
  var copy = DriveApp.getFileById(docId).makeCopy('pdf-export-tmp-' + Date.now());
  try {
    return exportPdf_(copy.getId());
  } finally {
    try { copy.setTrashed(true); } catch (e) {}
  }
}

/** md5 of a blob's bytes — fingerprint used to tell renders apart. */
function blobHash_(blob) {
  return Utilities.base64Encode(Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5, blob.getBytes()));
}

/** Queue a background re-export for an archived PDF that was a stale render. */
function enqueuePdfRefresh_(item) {
  var q = [];
  try { q = JSON.parse(PROPS.getProperty('PDF_REFRESH_Q') || '[]'); } catch (e) {}
  q.push(item);
  PROPS.setProperty('PDF_REFRESH_Q', JSON.stringify(q));
  ScriptApp.newTrigger('processPdfRefreshQueue_').timeBased()
    .after(3 * 60 * 1000).create();                    // first try ~3 min out
}

/** Time-triggered: re-export docs whose archived PDF was a stale render and
 *  overwrite the archived file with the real (post-edit) render. Retries
 *  every ~3 min, gives up after 6 attempts (~20 min). */
function processPdfRefreshQueue_() {
  var q = [];
  try { q = JSON.parse(PROPS.getProperty('PDF_REFRESH_Q') || '[]'); } catch (e) {}
  if (!q.length) return;
  var remaining = [];
  for (var i = 0; i < q.length; i++) {
    var it = q[i], done = false, gone = false;
    try {
      var blob = exportPdf_(it.docId);
      if (blobHash_(blob) !== it.baseline) {
        blob.setName(it.fname);
        // File.setBlob unavailable on these files — replace via create+trash
        var f = DriveApp.getFileById(it.fileId);
        var parent = null, pit = f.getParents();
        if (pit.hasNext()) parent = pit.next();
        if (parent) parent.createFile(blob); else DriveApp.getRootFolder().createFile(blob);
        f.setTrashed(true);
        done = true;
      }
    } catch (err) {
      // archived file already gone (user deleted it) -> nothing to heal
      if (String(err && err.message || err).toLowerCase().indexOf('not exist') >= 0 ||
          String(err && err.message || err).toLowerCase().indexOf('does not exist') >= 0) gone = true;
    }
    if (!done && !gone) {
      it.attempts = (it.attempts || 0) + 1;
      if (it.attempts <= 6) remaining.push(it);
    }
  }
  PROPS.setProperty('PDF_REFRESH_Q', JSON.stringify(remaining));
  if (remaining.length)
    ScriptApp.newTrigger('processPdfRefreshQueue_').timeBased()
      .after(3 * 60 * 1000).create();
}

/** refresh_pdf: force a fresh export of a customer's doc and overwrite the
 *  archived PDF for one invoice. Manual heal for files archived before the
 *  self-heal existed. body: { customer, invoice_no } */
function refreshPdf_(body) {
  var cust = String(body.customer || '').trim();
  var inv = String(body.invoice_no || '').trim();
  if (!cust || !inv) throw new Error('customer and invoice_no are required');
  var docId = customerDocId_(cust);
  if (!docId) throw new Error('no Google Doc mapped for ' + cust);
  var folder = customerFolder_(cust, false);
  if (!folder) throw new Error('no archived PDFs folder for ' + cust);
  var prefix = sanitizeFile_(inv), found = null;
  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (String(f.getName()).indexOf(prefix) === 0) { found = f; break; }
  }
  if (!found) throw new Error('no archived PDF for ' + inv + ' in BILLS/' + cust);
  var before = found.getBlob().getBytes().length;
  // copy-trick export: immune to the stale render cache
  var blob = exportFreshViaCopy_(docId);
  blob.setName(found.getName());
  // NOTE: File.setBlob proved unavailable on these files ('not a function'),
  // so overwrite = create the fresh file in the same folder, then trash old
  var parent = null, pit = found.getParents();
  if (pit.hasNext()) parent = pit.next();
  var nf = parent ? parent.createFile(blob) : DriveApp.getRootFolder().createFile(blob);
  found.setTrashed(true);
  return { ok: true, customer: cust, invoice_no: inv, pdf_name: found.getName(),
           old_bytes: before, new_bytes: blob.getBytes().length,
           changed: before !== blob.getBytes().length };
}

/* ------------------------------------------------------- edit an old bill - */

/** Locate ONE GST RECORDS row: invoice no + best fuzzy customer match
 *  (invoice numbers are not unique across customers). Returns the row index
 *  into gstVals (0-based) or -1. */
function findGstRow_(gstVals, cust, invNo) {
  var iName = colOf_(gstVals[0], 'Customer Name');
  var row = -1, best = 0;
  for (var r = 1; r < gstVals.length; r++) {
    if (String(gstVals[r][0] || '').trim() !== invNo) continue;
    var s = nameScore_(cust, String(gstVals[r][iName] || ''));
    if (s > best) { best = s; row = r; }
  }
  return row;
}

/** 'yyyy-mm-dd' -> Date (start of day) or null. */
function parseISO_(s) {
  var m = String(s || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

/** whole months between two dates s..e when e = s + m months - 1, else 0. */
function monthsBetween_(s, e) {
  if (!s || !e) return 0;
  for (var m = 1; m <= 120; m++) {
    if (addDays_(addMonths_(s, m), -1).getTime() === e.getTime()) return m;
  }
  return 0;
}

/** The plan-description block printed on the doc (the shouty lines, e.g.
 *  "HIGH SPEED 200 MBPS\nTRULY UNLIMITED 12 MONTHS."). Used as the bill-form
 *  prefill (exact PDF text) and as the old-value anchor when editing. */
function docDescBlock_(text) {
  var lines = String(text || '').split('\n');
  var best = [], cur = [];
  for (var i = 0; i <= lines.length; i++) {
    var s = i < lines.length ? lines[i].trim() : '';
    var ok = s.length >= 8 && !/[a-z]/.test(s) && /[A-Z]/.test(s) &&
             /MBPS|PLAN|SPEED|UNLIMITED|BROADBAND|FIBER|FIBRE/.test(s);
    if (ok) cur.push(s);
    else { if (cur.length > best.length) best = cur; cur = []; }
  }
  return best.join('\n');
}

/** bill_prefill: the Service Description EXACTLY as the customer's doc (and
 *  therefore their PDFs) show it — for the renewal bill form.
 *  body: { customer } */
function billPrefill_(body) {
  var cust = String(body.customer || '').trim();
  var desc = '', userid = '';
  var docId = customerDocId_(cust);
  if (docId) {
    try {
      desc = docDescBlock_(docText_(docId));
      userid = userIdInDoc_(docId);
    } catch (e) { /* doc read failed; caller retains any legacy fallback */ }
  }
  return { ok: true, customer: cust, desc: desc, userid: userid };
}

/** bill_edit_prefill: everything the edit-bill form needs for ONE existing
 *  bill. body: { customer, invoice_no } */
/**
 * Convert a printed invoice date (dd/mm/yy, dd-mm-yyyy, dd_mm_yy) to the
 * date input's yyyy-mm-dd value. Returns '' for malformed dates.
 */
function invoiceDateTextToISO_(value) {
  var m = String(value || '').trim().match(/^(\d{1,2})[\/._-](\d{1,2})[\/._-](\d{2,4})$/);
  if (!m) return '';
  var day = Number(m[1]), month = Number(m[2]), year = Number(m[3]);
  if (year < 100) year += 2000;
  var d = new Date(year, month - 1, day);
  if (d.getFullYear() !== year || d.getMonth() !== month - 1 || d.getDate() !== day) return '';
  return iso_(d);
}

/**
 * Older invoices live in the archive rather than the current customer Doc.
 * Generated archive filenames include the invoice number and its date.
 * Accept both hyphen- and underscore-separated dates because older edit
 * versions produced both naming conventions.
 */
function archivedInvoiceDateISO_(customer, invoiceNo) {
  var folder = customerFolder_(customer, false);
  if (!folder) return '';
  var prefix = sanitizeFile_(invoiceNo) + '_';
  var files = folder.getFiles(), bestISO = '', bestCreated = -1;
  while (files.hasNext()) {
    var file = files.next();
    var fileName = String(file.getName() || '');
    if (fileName.indexOf(prefix) !== 0) continue;
    var m = fileName.match(/_(\d{1,2})[-_](\d{1,2})[-_](\d{2,4})\.pdf$/i);
    if (!m) continue;
    var parsed = invoiceDateTextToISO_(m[1] + '-' + m[2] + '-' + m[3]);
    var created = file.getDateCreated().getTime();
    if (parsed && created >= bestCreated) {
      bestISO = parsed;
      bestCreated = created;
    }
  }
  return bestISO;
}

/**
 * Read the printed broadband User ID from the live Google Doc.
 * Supports "User id :- VALUE" in one cell or a label/value pair in adjacent
 * cells, and searches nested tables plus the document header/footer.
 */
function userIdInDoc_(docId) {
  var doc = DocumentApp.openById(docId);
  var roots = [doc.getBody(), doc.getHeader(), doc.getFooter()];

  function validValue_(value) {
    var v = String(value || '').replace(/\u00A0/g, ' ').trim();
    v = v.replace(/^[:\-]+\s*/, '').trim();
    if (!v) return '';
    if (/^(GSTIN|GSTIN \/ UIN|State Name|SAC Code|Invoice Date|Invoice No|Buyer\s*\(|Support|PAN|E-Mail|Place of Supply|State Code|User\s*id)\b/i.test(v)) return '';
    return v.split(/\r?\n/)[0].trim();
  }

  function valueFromLabelCell_(value) {
    var lines = String(value || '').replace(/\u00A0/g, ' ').split(/\r?\n/);
    for (var i = 0; i < lines.length; i++) {
      var match = lines[i].match(/User\s*id\s*(?:[:\-]+\s*)?(.*)$/i);
      if (!match) continue;
      var inlineValue = validValue_(match[1]);
      if (inlineValue) return inlineValue;
      for (var j = i + 1; j < lines.length; j++) {
        var nextLine = validValue_(lines[j]);
        if (nextLine) return nextLine;
      }
    }
    return '';
  }

  function scan_(element) {
    if (!element || typeof element.getType !== 'function') return '';
    if (element.getType() === DocumentApp.ElementType.TABLE) {
      var table = element;
      for (var r = 0; r < table.getNumRows(); r++) {
        var row = table.getRow(r), cells = [];
        for (var c = 0; c < row.getNumCells(); c++) cells.push(row.getCell(c));

        for (var c0 = 0; c0 < cells.length; c0++) {
          var labelText = cells[c0].getText();
          if (!/User\s*id/i.test(labelText)) continue;

          var sameCellValue = valueFromLabelCell_(labelText);
          if (sameCellValue) return sameCellValue;

          for (var c1 = c0 + 1; c1 < cells.length; c1++) {
            var adjacentValue = validValue_(cells[c1].getText());
            if (adjacentValue) return adjacentValue;
          }

          for (var r2 = r + 1; r2 < Math.min(table.getNumRows(), r + 3); r2++) {
            var belowRow = table.getRow(r2);
            if (belowRow.getNumCells() > c0) {
              var belowValue = validValue_(belowRow.getCell(c0).getText());
              if (belowValue) return belowValue;
            }
          }
        }

        for (var c2 = 0; c2 < cells.length; c2++) {
          var nestedValue = scan_(cells[c2]);
          if (nestedValue) return nestedValue;
        }
      }
      return '';
    }
    if (typeof element.getNumChildren === 'function') {
      for (var i = 0; i < element.getNumChildren(); i++) {
        var found = scan_(element.getChild(i));
        if (found) return found;
      }
    }
    return '';
  }

  for (var i0 = 0; i0 < roots.length; i0++) {
    if (!roots[i0]) continue;
    var result = scan_(roots[i0]);
    if (result) return result;
  }

  var flat = docText_(docId);
  var fallback = flat.match(/User\s*id\s*(?:[:\-]+\s*)?([^\r\n]+)/i);
  return fallback ? validValue_(fallback[1]) : '';
}

function billEditPrefill_(body) {
  var cust = String(body.customer || '').trim();
  var inv = String(body.invoice_no || '').trim();
  if (!cust || !inv) throw new Error('customer and invoice_no are required');
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var gst = ss.getSheetByName(GST_SHEET);
  var up = ss.getSheetByName(UPCOMING_SHEET);
  var gstVals = gst.getDataRange().getValues();
  var gCols = {};
  GST_COLUMNS.forEach(function (c) { gCols[c] = colOf_(gstVals[0], c); });
  var row = findGstRow_(gstVals, cust, inv);
  if (row < 0) throw new Error('no GST RECORDS row with invoice ' + inv + ' for ' + cust);
  var dParsed = gCols['Invoice Date'] >= 0 ? asDate_(gstVals[row][gCols['Invoice Date']]) : null;
  var rec = {
    name: String(gstVals[row][gCols['Customer Name']] || '').trim(),
    gstin: gstinDocumentValue_(gstVals[row][gCols['Customer GSTIN']]),
    desc: String(gstVals[row][gCols['Service Description']] || '').trim(),
    amount: num_(gstVals[row][gCols['Amount ₹']]),
    date: dParsed ? iso_(dParsed) : ''
  };

  // UPCOMING BILLS row: address / user id / expiry / duration live there
  var upVals = up.getDataRange().getValues();
  var iName = colOf_(upVals[0], 'Customer name'), iExp = colOf_(upVals[0], 'Expire Date');
  var iDur = colOf_(upVals[0], COL_DURATION), iAddr = colOf_(upVals[0], COL_ADDRESS);
  var iUser = colOf_(upVals[0], COL_USERID);
  var upRow = -1, best = 0;
  for (var r = 1; r < upVals.length; r++) {
    var s = nameScore_(cust, String(upVals[r][iName] || ''));
    if (s > best) { best = s; upRow = r; }
  }
  var address = upRow >= 0 && iAddr >= 0 ? String(upVals[upRow][iAddr] || '').trim() : '';
  var userid = upRow >= 0 && iUser >= 0 ? String(upVals[upRow][iUser] || '').trim() : '';
  var months = upRow >= 0 && iDur >= 0 ? parseDuration_(String(upVals[upRow][iDur] || '')) : 0;
  if (!(months > 0)) {
    var dm = rec.desc.match(/(\d+)\s*month/i);
    if (dm) months = parseInt(dm[1], 10);
  }

  // is this the bill the doc currently shows? (doc holds the LATEST bill)
  var docId = customerDocId_(cust) || customerDocId_(rec.name);
  var docText = '';
  try { if (docId) docText = docText_(docId); } catch (e) { /* ignore */ }
  var isLatest = !!docText && docText.indexOf(inv) >= 0;
  var prefillBuyerName = rec.name;
  var prefillUserId = userid;
  if (isLatest && docId) {
    try {
      var actualBuyerName = buyerNameInDoc_(docId);
      if (actualBuyerName) prefillBuyerName = actualBuyerName;
    } catch (buyerReadErr) { /* retain GST RECORDS fallback */ }
  }
  // The customer Doc is the source document used for both latest and older
  // invoice PDF edit copies, so use its printed User ID for either invoice.
  if (docId) {
    try {
      var actualUserId = userIdInDoc_(docId);
      if (actualUserId) prefillUserId = actualUserId;
    } catch (userReadErr) { /* retain legacy UPCOMING BILLS fallback */ }
  }

  // Prefill from the bill's printed source rather than blindly trusting a
  // sheet date that may have been changed by an earlier buggy edit.
  var invoiceDate = rec.date;
  var invoiceDateSource = 'GST RECORDS (fallback)';
  if (isLatest && docId) {
    try {
      var printedDate = invoiceDateTextInDoc_(docId);
      var printedISO = invoiceDateTextToISO_(printedDate);
      if (printedISO) {
        invoiceDate = printedISO;
        invoiceDateSource = 'Google Doc';
      }
    } catch (dateReadErr) { /* retain the GST RECORDS fallback */ }
  } else {
    var archivedISO = archivedInvoiceDateISO_(rec.name || cust, inv);
    if (archivedISO) {
      invoiceDate = archivedISO;
      invoiceDateSource = 'archived PDF filename (verify against PDF)';
    }
  }

  // period prefill: 1) the hidden "Bill Period" sheet column (authoritative
  // track) 2) the doc's own period line (latest bill) 3) derived guess
  var ps = null, pe = null;
  var pCell = String(gstVals[row][ensureBillPeriodCol_(gst)] || '');
  var pM = pCell.match(/(\d{2}\/\d{2}\/\d{4}) to (\d{2}\/\d{2}\/\d{4})/);
  if (pM) { ps = dmyToISO_(pM[1]); pe = dmyToISO_(pM[2]); }
  var perM = docText.match(/\d{2}\/\d{2}\/\d{4} to \d{2}\/\d{2}\/\d{4}/);
  if (!ps && isLatest && perM) {
    var pp = perM[0].split(' to ');
    ps = dmyToISO_(pp[0]); pe = dmyToISO_(pp[1]);
  } else if (!ps && dParsed) {
    ps = rec.date;
    if (months > 0) pe = iso_(addDays_(addMonths_(dParsed, months), -1));
  }
  return {
    ok: true, customer: cust, invoice_no: inv, is_latest: isLatest,
    name: prefillBuyerName, gstin: rec.gstin, userid: prefillUserId, address: address,
    invoice_date: invoiceDate, invoice_date_source: invoiceDateSource,
    amount: rec.amount,
    desc: (isLatest && docDescBlock_(docText)) ? docDescBlock_(docText) : rec.desc,
    doc_desc: docDescBlock_(docText),
    period_start: ps || '', period_end: pe || '', months: months
  };
}

/** backfill_periods: one-time heal — write each mapped customer's doc period
 *  line into the hidden "Bill Period" column on that bill's GST row. Only
 *  the LATEST bill per customer can be filled this way (the doc shows only
 *  it); older rows fill as bills get edited from the app. */
function backfillPeriods_(body) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var gst = ss.getSheetByName(GST_SHEET);
  var map;
  try { map = JSON.parse(PROPS.getProperty('DOCS_MAP') || '{}'); } catch (e) { map = {}; }
  var c0 = ensureBillPeriodCol_(gst);
  var filled = 0, skipped = [];
  var gstVals = gst.getDataRange().getValues();
  var iNo = colOf_(gstVals[0], 'Invoice No');
  for (var cust in map) {
    var text = '';
    try { text = docText_(map[cust]); } catch (e) { skipped.push(cust + ': doc unread'); continue; }
    var perM = text.match(/\d{2}\/\d{2}\/\d{4} to \d{2}\/\d{2}\/\d{4}/);
    var invM = text.match(/[A-Z]{2,5}\/\d{2,4}\/\d{2,4}-\d{2,4}/) || text.match(/ST\/\d+\/[\d-]+/);
    if (!perM || !invM) { skipped.push(cust + ': no period/invoice line in doc'); continue; }
    var wrote = false;
    for (var r = 1; r < gstVals.length; r++) {
      if (String(gstVals[r][iNo] || '').trim() === invM[0]) {
        gst.getRange(r + 1, c0 + 1).setValue(perM[0]);
        gstVals[r][c0] = perM[0];
        filled++; wrote = true;
        break;
      }
    }
    if (!wrote) skipped.push(cust + ': ' + invM[0] + ' not found in GST RECORDS');
  }
  return { ok: true, filled: filled, skipped: skipped };
}

/** 'dd/mm/yyyy' -> 'yyyy-mm-dd' */
function dmyToISO_(s) {
  var m = String(s || '').trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return '';
  var d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  return iso_(d);
}

/** edit_bill: edit an EXISTING bill end-to-end.
 *  Editable: buyer name, address, GSTIN/UIN, invoice date, user id, bill
 *  period, service description, amount (CGST/SGST/total/words recompute).
 *  NOT editable: invoice no (later bills chain off it).
 *  RENAMING touches ONLY the doc/PDF — never the sheet's Customer Name.
 *  Pipeline (design by user):
 *  - LATEST bill: edit the customer's doc IN PLACE (normal flow), export
 *    fresh (copy-trick), replace that bill's archived PDF; UPCOMING BILLS
 *    expiry/duration follow the new period end.
 *  - OLDER bill: the doc is NEVER touched — a throwaway COPY of the doc is
 *    edited to the old bill's corrected values, exported (fresh render) and
 *    trashed; only the archived PDF for that invoice is replaced. The doc
 *    keeps showing the latest bill the whole time.
 *  Sheet: the bill's own GST RECORDS row is updated in both cases. */
function editBill_(body) {
  var inv = String(body.invoice_no || '').trim();
  var cust = String(body.customer || '').trim();
  if (!cust || !inv) throw new Error('customer and invoice_no are required');
  var amount = parseFloat(body.amount);
  if (!(amount > 0)) throw new Error('amount is required');
  var cgst = round2_(amount * 0.09), sgst = round2_(amount * 0.09);
  var total = round2_(amount + cgst + sgst);
  var desc = String(body.desc || '').trim();
  var name = String(body.name || '').trim();
  var address = String(body.address || '').trim();
  var userid = String(body.userid || '').trim();
  var gstin = gstinDocumentValue_(body.gstin);
  // Existing bill edits preserve the invoice date selected in the edit form.
  // An edit must never silently replace the original date with today's date.
  var requestedDateText = String(body.invoice_date || '').trim();
  var newDate = requestedDateText ? parseISO_(requestedDateText) : null;
  if (requestedDateText && !newDate)
    throw new Error('invoice_date must be a valid yyyy-mm-dd date');
  var ps = parseISO_(body.period_start), pe = parseISO_(body.period_end);
  if (!(ps && pe)) throw new Error('bill period (start + end) is required — ' +
    'read it from the bill\'s current PDF (View in the app)');

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var gst = ss.getSheetByName(GST_SHEET);
  var up = ss.getSheetByName(UPCOMING_SHEET);
  var gstVals = gst.getDataRange().getValues();
  var gCols = {};
  GST_COLUMNS.forEach(function (c) { gCols[c] = colOf_(gstVals[0], c); });

  // --- 1. the bill's GST row ------------------------------------------------
  var row = findGstRow_(gstVals, cust, inv);
  if (row < 0) throw new Error('no GST RECORDS row with invoice ' + inv + ' for ' + cust);
  var billCust = String(gstVals[row][gCols['Customer Name']] || '').trim();
  var updateGstName = body.update_gst_name === true;
  var updateUpcomingName = body.update_upcoming_name === true;
  if ((updateGstName || updateUpcomingName) && !name)
    throw new Error('Buyer name is required when a sheet rename option is selected');

  var upValsBefore = up.getDataRange().getValues();
  var iUpcomingName = colOf_(upValsBefore[0], 'Customer name');
  var upcomingNameRow = -1, upcomingNameScore = 0;
  var oldUpcomingName = '';
  if (updateUpcomingName) {
    if (iUpcomingName < 0) throw new Error('UPCOMING BILLS has no Customer name column');
    for (var ur = 1; ur < upValsBefore.length; ur++) {
      var us = nameScore_(cust, String(upValsBefore[ur][iUpcomingName] || ''));
      if (us > upcomingNameScore) {
        upcomingNameScore = us;
        upcomingNameRow = ur;
      }
    }
    if (upcomingNameRow < 1 || upcomingNameScore < 0.8)
      throw new Error('Could not uniquely locate the customer in UPCOMING BILLS; no changes were saved');
    oldUpcomingName = String(upValsBefore[upcomingNameRow][iUpcomingName] || '').trim();

    for (var ux = 1; ux < upValsBefore.length; ux++) {
      if (ux === upcomingNameRow) continue;
      var otherName = String(upValsBefore[ux][iUpcomingName] || '').trim();
      if (otherName && normName_(otherName) === normName_(name))
        throw new Error('UPCOMING BILLS already contains another customer named "' + name + '". Choose a different name; no changes were saved.');
    }
  }

  // --- 2. the customer's doc ------------------------------------------------
  var docId = customerDocId_(cust) || customerDocId_(billCust);
  if (!docId) throw new Error('no Google Doc mapped for ' + cust + '/' + billCust);
  var before = docText_(docId);
  var docDescription = desc || docDescBlock_(before);
  var sheetDescription = conciseServiceDescription_(docDescription);
  if (!sheetDescription) {
    sheetDescription = previousServiceDescription_(gstVals, cust, inv) ||
      String(gstVals[row][gCols['Service Description']] || '').trim();
  }
  var isLatest = before.indexOf(inv) >= 0;   // doc currently shows THIS bill
  if (!isLatest && !(ps && pe)) {
    throw new Error('this is not the latest bill — the Bill Period (start + end) is ' +
      'required so the doc can be temporarily rewritten and restored');
  }

  // --- 3. the bill the doc CURRENTLY shows (old-value anchors + restore) ----
  var docInvM = before.match(/[A-Z]{2,5}\/\d{2,4}\/\d{2,4}-\d{2,4}/) ||
                before.match(/ST\/\d+\/[\d-]+/);
  var docBuyerName = isLatest ? buyerNameInDoc_(docId) : '';
  // Read the invoice date from the document itself. The GST RECORDS row may
  // already contain today's date from an earlier edit while the Doc still
  // displays its previous date.
  var docDateText = invoiceDateTextInDoc_(docId);
  var docDateM = docDateText ? [docDateText, docDateText] : null;

  // Defensive fallback for older clients that omit invoice_date:
  // latest bill -> live Google Doc; older bill -> archived PDF filename;
  // only then fall back to GST RECORDS.
  if (!newDate) {
    var fallbackISO = isLatest ? invoiceDateTextToISO_(docDateText) :
      archivedInvoiceDateISO_(billCust, inv);
    if (!fallbackISO && gCols['Invoice Date'] >= 0) {
      var sheetDate = asDate_(gstVals[row][gCols['Invoice Date']]);
      if (sheetDate) fallbackISO = iso_(sheetDate);
    }
    newDate = fallbackISO ? parseISO_(fallbackISO) : null;
  }
  if (!newDate)
    throw new Error('Could not determine this bill\'s existing invoice date. Reopen Edit Bill and select the correct invoice date; no changes were saved.');

  var docRow = null;
  if (docInvM) {
    for (var r2 = 1; r2 < gstVals.length; r2++) {
      if (String(gstVals[r2][0] || '').trim() === docInvM[0]) {
        docRow = {
          invoiceNo: docInvM[0],
          invoiceDateText: docDateM ? docDateM[1] : '',
          customer: docBuyerName || String(gstVals[r2][gCols['Customer Name']] || '').trim(),
          gstin: gstinDocumentValue_(gstVals[r2][gCols['Customer GSTIN']]),
          amount: num_(gstVals[r2][gCols['Amount ₹']]),
          cgst: num_(gstVals[r2][gCols['CGST 9% (₹)']]),
          sgst: num_(gstVals[r2][gCols['SGST 9% (₹)']]),
          total: num_(gstVals[r2][gCols['Total Amount (₹)']])
        };
        break;
      }
    }
  }
  if (!docRow) throw new Error('could not identify the bill the doc currently shows (' +
    (docInvM ? docInvM[0] : 'no invoice no found in doc') + ') — edit aborted');
  if (!docDateM) throw new Error('could not find the Invoice Date printed in the customer Google Doc — edit aborted without changing the bill');

  // Invoice numbers use ST/<serial>/<FY>. Keep the serial for corrections
  // within the same FY; allocate the next unused serial if the edited date
  // moves this invoice into a different financial year.
  var newInv = inv;
  var invParts = inv.match(INVOICE_RE);
  if (invParts && (invParts[2] + '-' + invParts[3]) !== fyOf_(newDate)) {
    newInv = nextInvoiceNo_(gstVals, newDate);
    for (var ir = 1; ir < gstVals.length; ir++) {
      if (String(gstVals[ir][gCols['Invoice No']] || '').trim() === newInv &&
          ir !== row) {
        throw new Error('generated invoice number already exists: ' + newInv +
          ' — no changes were saved');
      }
    }
  }

  // --- 4. build old -> new pairs anchored on the doc's ACTUAL text ----------
  var perM = before.match(/\d{2}\/\d{2}\/\d{4} to \d{2}\/\d{2}\/\d{4}/);
  var wordM = before.match(/[A-Z][A-Z ,\-]*RUPEES ONLY[^\n]*/);
  function words(n) { return rupeeWords_(n) + ' RUPEES ONLY/_'; }
  var pairs = [
    { name: 'Invoice No',   old: docRow.invoiceNo,          nw: newInv },
    { name: 'Amount',       old: money_(docRow.amount),     nw: money_(amount) },
    { name: 'GST cells',    old: money_(docRow.cgst),       nw: money_(cgst) },
    { name: 'Total',        old: money_(docRow.total),      nw: money_(total) },
    { name: 'Words',        old: wordM ? wordM[0] : words(docRow.total), nw: words(total) }
  ];
  if (newDate && docRow.invoiceDateText)
    pairs.push({ name: 'Invoice Date', old: docRow.invoiceDateText,
                 nw: formatInvoiceDateLike_(newDate, docRow.invoiceDateText) });
  if (ps && pe) pairs.push({ name: 'Bill Period',
    old: perM ? perM[0] : '', nw: fmtDMY_(ps) + ' to ' + fmtDMY_(pe) });
  if (desc) {
    var oldBlock = docDescBlock_(before);
    pairs.push({ name: 'Description', old: oldBlock || String(docRow.desc || ''), nw: desc, warn: true });
  }
  // buyer / address / gstin / user id — anchored on the doc's actual text
  function locate_(needle) {
    needle = String(needle || '');
    if (!needle) return '';
    if (before.indexOf(needle) >= 0) return needle;
    var j = before.toLowerCase().indexOf(needle.toLowerCase());
    return j >= 0 ? before.substr(j, needle.length) : '';
  }
  var buyerOld = locate_(docRow.customer);
  if (buyerOld && name && buyerOld !== name) pairs.push({ name: 'Buyer', old: buyerOld, nw: name, warn: true });
  var gstinOld = locate_(docRow.gstin);
  if (gstinOld && gstin && gstinOld !== gstin) pairs.push({ name: 'GSTIN', old: gstinOld, nw: gstin, warn: true });
  if (address || userid) {
    var upVals = up.getDataRange().getValues();
    var iName = colOf_(upVals[0], 'Customer name');
    var iAddr = colOf_(upVals[0], COL_ADDRESS), iUser = colOf_(upVals[0], COL_USERID);
    var upRow = -1, best = 0;
    for (var r3 = 1; r3 < upVals.length; r3++) {
      var sc = nameScore_(cust, String(upVals[r3][iName] || ''));
      if (sc > best) { best = sc; upRow = r3; }
    }
    if (upRow >= 0) {
      var addrOld = iAddr >= 0 ? locate_(String(upVals[upRow][iAddr] || '').trim()) : '';
      if (addrOld && address && addrOld !== address) pairs.push({ name: 'Address', old: addrOld, nw: address, warn: true });
      // For latest bills this anchor updates the live Doc; for older bills
      // editCustomerDoc_ applies it only to the throwaway PDF-generation copy.
      var userOld = userIdInDoc_(docId);
      if (!userOld && iUser >= 0)
        userOld = locate_(String(upVals[upRow][iUser] || '').trim());
      if (userOld && userid && userOld !== userid)
        pairs.push({ name: 'User ID', old: userOld, nw: userid, warn: true });
    }
  }

  // --- 5. render the edited bill's PDF --------------------------------------
  // LATEST bill: edit the real doc in place, then export via copy-trick.
  // OLDER bill: edit a THROWAWAY COPY of the doc and export that — the
  // customer's doc is never touched, so the latest bill stays intact.
  var edit, blob;
  if (isLatest) {
    edit = editCustomerDoc_(docId, pairs);
    // replaceTextInDoc_ flushes the exact document session after each write.
    Utilities.sleep(1000);
    var latestPrintedDate = invoiceDateTextInDoc_(docId);
    var latestExpectedDate = formatInvoiceDateLike_(newDate, docDateText);
    if (latestPrintedDate !== latestExpectedDate) {
      throw new Error('Invoice Date verification failed. Expected ' + latestExpectedDate +
        ' in the Google Doc, but found ' + (latestPrintedDate || '(not found)') +
        '. PDF and sheet were not updated.');
    }
    blob = exportFreshViaCopy_(docId);
  } else {
    var tmp = DriveApp.getFileById(docId).makeCopy('bill-edit-tmp-' + Date.now());
    try {
      Utilities.sleep(1500);                       // let the fresh copy settle
      edit = editCustomerDoc_(tmp.getId(), pairs);
      Utilities.sleep(1000);                       // allow the saved copy to render
      var olderPrintedDate = invoiceDateTextInDoc_(tmp.getId());
      var olderExpectedDate = formatInvoiceDateLike_(newDate, docDateText);
      if (olderPrintedDate !== olderExpectedDate) {
        throw new Error('Invoice Date verification failed on the temporary bill copy. Expected ' +
          olderExpectedDate + ', found ' + (olderPrintedDate || '(not found)') +
          '. PDF was not generated and GST RECORDS was not updated.');
      }
      blob = exportFreshViaCopy_(tmp.getId());
    } finally {
      try { tmp.setTrashed(true); } catch (e2) { /* never block on cleanup */ }
    }
  }

  // --- 6. replace that bill's archived PDF ---------------------------------
  var dateStr = fmtShort_(newDate);                // dd/mm/yy
  var fname = sanitizeFile_(newInv) + '_' + sanitizeFile_(dateStr).replace(/\//g, '-') + '.pdf';
  blob.setName(fname);
  var folder = customerFolder_(cust, true);
  var prefix = sanitizeFile_(inv), oldFiles = [], it = folder.getFiles();
  while (it.hasNext()) {
    var f0 = it.next();
    if (String(f0.getName()).indexOf(prefix) === 0) oldFiles.push(f0);
  }
  // Create the verified replacement first, then remove every active stale
  // copy for this invoice (older runs may have left duplicate PDFs behind).
  folder.createFile(blob);
  oldFiles.forEach(function (oldFile) { oldFile.setTrashed(true); });

  // --- 8. sheet: same GST RECORDS row (Customer Name is NEVER changed here —
  //         renaming is doc/PDF only, per design) ----------------------------
  var gstNow = gst.getDataRange().getValues();      // re-read: rows may have shifted
  var rowNow = findGstRow_(gstNow, cust, inv);
  if (rowNow < 0) throw new Error('GST RECORDS row vanished while editing — re-check the sheet');
  var gInv = colOf_(gstNow[0], 'Invoice No');
  var gGstin = colOf_(gstNow[0], 'Customer GSTIN');
  var gDesc = colOf_(gstNow[0], 'Service Description'), gAmt = colOf_(gstNow[0], 'Amount ₹');
  var gCg = colOf_(gstNow[0], 'CGST 9% (₹)'), gSg = colOf_(gstNow[0], 'SGST 9% (₹)');
  var gTot = colOf_(gstNow[0], 'Total Amount (₹)'), gDt = colOf_(gstNow[0], 'Invoice Date');
  var gCustomerName = colOf_(gstNow[0], 'Customer Name');
  if (newInv !== inv && gInv >= 0)
    gst.getRange(rowNow + 1, gInv + 1).setValue(newInv);
  if (gGstin >= 0) gst.getRange(rowNow + 1, gGstin + 1).setValue(gstinSheetValue_(gstin));
  if (gDesc >= 0 && sheetDescription)
    gst.getRange(rowNow + 1, gDesc + 1).setValue(sheetDescription);
  if (gAmt >= 0) gst.getRange(rowNow + 1, gAmt + 1).setValue(amount);
  if (gCg >= 0) gst.getRange(rowNow + 1, gCg + 1).setValue(cgst);
  if (gSg >= 0) gst.getRange(rowNow + 1, gSg + 1).setValue(sgst);
  if (gTot >= 0) gst.getRange(rowNow + 1, gTot + 1).setValue(total);
  if (newDate && gDt >= 0) {
    var dCell = gst.getRange(rowNow + 1, gDt + 1);
    dCell.setNumberFormat('dd/mm/yy');
    dCell.setValue(newDate);
  }
  if (ps && pe) writeBillPeriod_(gst, ensureBillPeriodCol_(gst), rowNow, ps, pe);

  var gstNameUpdated = false;
  if (updateGstName && gCustomerName >= 0) {
    var changedGstRows = 0;
    for (var gr = 1; gr < gstNow.length; gr++) {
      if (String(gstNow[gr][gCustomerName] || '').trim() === billCust) {
        gst.getRange(gr + 1, gCustomerName + 1).setValue(name);
        changedGstRows++;
      }
    }
    gstNameUpdated = changedGstRows > 0;
  }

  var upcomingNameUpdated = false;
  if (updateUpcomingName && upcomingNameRow >= 1 && iUpcomingName >= 0) {
    up.getRange(upcomingNameRow + 1, iUpcomingName + 1).setValue(name);
    upcomingNameUpdated = true;
  }

  if ((gstNameUpdated || upcomingNameUpdated) && name) {
    rememberCustomerRenameAliases_([cust, billCust, oldUpcomingName], name, docId,
      updateGstName ? name : billCust);
  }

  // --- 9. latest bill only: UPCOMING BILLS expiry / duration follow --------
  var upUpdated = false;
  if (isLatest && ps && pe) {
    var months = monthsBetween_(ps, pe);
    var upVals2 = up.getDataRange().getValues();
    var iN2 = colOf_(upVals2[0], 'Customer name'), iE2 = colOf_(upVals2[0], 'Expire Date');
    var iD2 = colOf_(upVals2[0], COL_DURATION);
    for (var r4 = 1; r4 < upVals2.length; r4++) {
      var expiryLookupName = updateUpcomingName ? name : cust;
      if (nameScore_(expiryLookupName, String(upVals2[r4][iN2] || '')) >= 0.8) {
        if (iE2 >= 0) {
          var eCell = up.getRange(r4 + 1, iE2 + 1);
          eCell.setNumberFormat('dd/mm/yy');
          eCell.setValue(pe);
        }
        if (iD2 >= 0 && months > 0) up.getRange(r4 + 1, iD2 + 1).setValue(durationText_(months));
        upUpdated = true;
        break;
      }
    }
  }

  var currentUpcomingHeader = up.getRange(1, 1, 1, up.getLastColumn()).getValues()[0];
  var currentExpireIndex = colOf_(currentUpcomingHeader, 'Expire Date');
  if (currentExpireIndex >= 0) {
    // Re-sort the entire UPCOMING BILLS table after every bill edit. This
    // also repairs any rows that were already out of order before this edit.
    sortUpcomingByExpire_(up, currentExpireIndex);
  }

  return {
    ok: true, customer: cust, invoice_no: newInv,
    previous_invoice_no: newInv !== inv ? inv : null,
    invoice_no_changed: newInv !== inv, is_latest: isLatest,
    doc_restored: !isLatest, upcoming_updated: upUpdated,
    gst_name_updated: gstNameUpdated, upcoming_name_updated: upcomingNameUpdated,
    updated_customer_name: name,
    edited: edit.edited, skipped: edit.skipped, desc: edit.desc,
    amount: amount, cgst: cgst, sgst: sgst, total: total,
    pdf_name: fname, pdf_base64: Utilities.base64Encode(blob.getBytes())
  };
}

/** The new bill pipeline: edit doc in place -> export -> archive. */
function billViaCustomerDoc_(custName, last, oldExp, lastMonths, bill) {
  var docId = customerDocId_(custName);
  if (!docId) throw new Error('no Google Doc mapped for "' + custName +
    '" — run test_backend.py docs (sync) and check_docs first');
  var pairs = docEditPairs_(last, oldExp, lastMonths, bill);
  // Anchor the critical pairs on the doc's ACTUAL current text where the
  // back-computed values can drift: month-end periods (01/11->30/11 vs
  // back-computed 31/10->30/11) and words-line formatting. The doc's own
  // line is the truth for "old"; the new values come from the bill.
  var docText = docText_(docId);
  var perM = docText.match(/\d{2}\/\d{2}\/\d{4} to \d{2}\/\d{2}\/\d{4}/);
  var wordM = docText.match(/[A-Z][A-Z ,\-]*RUPEES ONLY[^\n]*/);
  var dateText = invoiceDateTextInDoc_(docId);
  if (!dateText) {
    throw new Error('Invoice Date label/value could not be located in the customer Google Doc table. The document was not changed and GST RECORDS was not updated.');
  }
  var currentDocUserId = userIdInDoc_(docId);
  var desiredUserId = String(bill.userid || currentDocUserId || '').trim();
  if (currentDocUserId && desiredUserId && currentDocUserId !== desiredUserId) {
    pairs.push({ name: 'User ID', old: currentDocUserId, nw: desiredUserId, warn: true });
  }
  pairs.forEach(function (p) {
    if (p.name === 'Bill Period' && perM && p.old !== p.nw) p.old = perM[0];
    if (p.name === 'Invoice Date') {
      p.old = dateText;
      p.nw = formatInvoiceDateLike_(bill.invoiceDate, dateText);
    }
    if (p.name === 'Words' && wordM && p.old !== p.nw) p.old = wordM[0];
  });
  // Fingerprint the PRE-EDIT render first (also warms the export cache) —
  // exportAndArchive_ uses it to tell the stale cached render from the real
  // post-edit one.
  var baselineHash = null;
  try { baselineHash = blobHash_(exportPdf_(docId)); } catch (e) { /* optional */ }
  var edit = editCustomerDoc_(docId, pairs);
  Utilities.sleep(300);
  var expectedPrintedDate = formatInvoiceDateLike_(bill.invoiceDate, dateText);
  var actualPrintedDate = invoiceDateTextInDoc_(docId);
  if (actualPrintedDate !== expectedPrintedDate) {
    throw new Error('Invoice Date verification failed. Expected ' + expectedPrintedDate +
      ' in the Google Doc, but found ' + (actualPrintedDate || '(not found)') +
      '. PDF generation and GST RECORDS update were stopped.');
  }
  var arch = exportAndArchive_(docId, custName, bill.invoiceNo,
                               String(bill.invoiceDate || '').replace(/\//g, '-'), baselineHash);
  return { blob: arch.blob, source: 'customer-doc', file: arch.file, folderUrl: arch.folderUrl,
           pdf_settled: arch.pdf_settled,
           edited: edit.edited, skipped: edit.skipped, desc: edit.desc };
}

/** add_customer: onboard a NEW customer end-to-end:
 *  1. creates their Google Doc by copying a mapped customer's doc (identical
 *     layout, pure DocumentApp — no template needed) and replacing the
 *     donor's values with the new customer's first-bill values
 *  2. exports + archives the first invoice PDF under BILLS/<Customer>/
 *  3. appends the first invoice to GST RECORDS (Payment Status: Paid)
 *  4. appends them to UPCOMING BILLS (expire = first period end) + re-sorts
 *  5. maps the new doc (DOCS_MAP)
 *  The sheet is written only after the doc + PDF succeeded; on any doc/parse
 *  failure the copy is trashed and nothing is saved. */
function addCustomer_(body) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Another customer or billing operation is running. Please retry in a moment.');
  }
  try {
    return addCustomerLocked_(body || {});
  } finally {
    lock.releaseLock();
  }
}


/**
 * Creates a new customer bill from the master Google Docs template.
 * No existing customer's document is used as a source.
 */
function addCustomerLocked_(body) {
  var name = String(body.name || '').trim();
  if (!name) throw new Error('Customer name is required.');
  var requestedSheetName = String(body.sheet_name || '').trim();
  if (requestedSheetName && requestedSheetName.toLowerCase() !== name.toLowerCase()) {
    throw new Error('For a mismatch-free onboarding, Customer Name, GST RECORDS name and UPCOMING BILLS name must be identical. Please use one customer name.');
  }
  var sheetName = name;
  var docName = sanitizeFile_(name);
  var address = String(body.address || '').replace(/\s*\r?\n\s*/g, ', ').trim();
  var userid = String(body.userid || '').trim();
  var amount = Number(body.amount);
  var months = parseInt(body.duration_months, 10);
  var desc = String(body.description || '').trim();
  var mode = String(body.payment_mode || '').trim() || 'Cash';
  var gstin = gstinDocumentValue_(body.gstin);

  if (!address) throw new Error('Customer address is required.');
  if (!userid) throw new Error('User ID is required.');
  if (!isFinite(amount) || amount <= 0) throw new Error('Enter a valid plan amount greater than zero.');
  if (!(months > 0)) throw new Error('Select a valid bill duration.');
  if (!desc) throw new Error('Service description is required.');
  if (!/^[\w\s.,#()/&+-]+$/.test(mode)) throw new Error('Payment mode contains unsupported characters.');

  var templateId = String(PROPS.getProperty('NEW_CUSTOMER_TEMPLATE_ID') || '').trim();
  if (!templateId) {
    throw new Error('The master onboarding template is not configured. Upload MASTER_CUSTOMER_ONBOARDING_TEMPLATE.docx to Drive, convert it to Google Docs, then save its Google Doc ID in Settings.');
  }
  var templateFile;
  try {
    templateFile = DriveApp.getFileById(templateId);
  } catch (templateError) {
    throw new Error('The configured master template could not be opened. Check NEW_CUSTOMER_TEMPLATE_ID and confirm the Apps Script owner has access.');
  }
  if (templateFile.getMimeType() !== MimeType.GOOGLE_DOCS) {
    throw new Error('The onboarding template must be a native Google Doc. Open the uploaded DOCX with Google Docs and use File → Save as Google Docs, then configure the new Google Doc ID.');
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var up = ss.getSheetByName(UPCOMING_SHEET);
  var gst = ss.getSheetByName(GST_SHEET);
  if (!up || !gst) throw new Error('Tabs "' + UPCOMING_SHEET + '" / "' + GST_SHEET + '" were not found.');

  var upVals = up.getDataRange().getValues();
  var upHeaders = upVals[0] || [];
  var iName = colOf_(upHeaders, 'Customer name');
  var iExp = colOf_(upHeaders, 'Expire Date');
  var iSrNo = upcomingSrNoColumn_(upHeaders);
  var iDur = colOf_(upHeaders, COL_DURATION);
  var iAddr = colOf_(upHeaders, COL_ADDRESS);
  var iUser = colOf_(upHeaders, COL_USERID);
  if (iName < 0 || iExp < 0) throw new Error('UPCOMING BILLS needs "Customer name" and "Expire Date" columns.');
  if (iSrNo < 0) throw new Error('UPCOMING BILLS needs an Sr No column (accepted headers: Sr No, Sr. No., S.No., Serial No).');

  var gstVals = gst.getDataRange().getValues();
  var gHeaders = gstVals[0] || [];
  var gCols = {};
  GST_COLUMNS.forEach(function (colName) {
    gCols[colName] = colOf_(gHeaders, colName);
    if (gCols[colName] < 0) throw new Error('GST RECORDS is missing the "' + colName + '" column.');
  });
  var gstNameCol = colOf_(gHeaders, 'Customer Name');
  var gstInvCol = colOf_(gHeaders, 'Invoice No');

  // Reject duplicates across both active data sources and the document map.
  for (var r = 1; r < upVals.length; r++) {
    if (String(upVals[r][iName] || '').trim().toLowerCase() === sheetName.toLowerCase()) {
      throw new Error('Customer already exists in UPCOMING BILLS: ' + sheetName);
    }
  }
  for (var gr = 1; gr < gstVals.length; gr++) {
    if (gstNameCol >= 0 &&
        String(gstVals[gr][gstNameCol] || '').trim().toLowerCase() === sheetName.toLowerCase()) {
      throw new Error('Customer already has an entry in GST RECORDS: ' + sheetName);
    }
  }

  var map;
  try { map = JSON.parse(PROPS.getProperty('DOCS_MAP') || '{}'); } catch (mapError) { map = {}; }
  Object.keys(map).forEach(function (mappedName) {
    if (mappedName.toLowerCase() === sheetName.toLowerCase() && map[mappedName]) {
      throw new Error('A Google Doc is already mapped to this customer: ' + mappedName);
    }
  });

  // Keep the document filename aligned to the single customer name source.
  var docsRoot = docsFolder_();
  var sameNameFiles = docsRoot.getFilesByName(docName);
  while (sameNameFiles.hasNext()) {
    var existingFile = sameNameFiles.next();
    if (existingFile.getMimeType() === MimeType.GOOGLE_DOCS) {
      throw new Error('A Google Doc named "' + docName + '" already exists in the bills folder. Rename or reconcile the existing document before onboarding.');
    }
  }

  // India-local invoice date; bill period start is the date selected by the user.
  var today = indiaToday_();
  var periodStart = today;
  var psRaw = String(body.period_start || '').trim();
  if (psRaw) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(psRaw)) throw new Error('Bill period start must be a valid date.');
    var q = psRaw.split('-');
    periodStart = new Date(Number(q[0]), Number(q[1]) - 1, Number(q[2]));
    if (periodStart.getFullYear() !== Number(q[0]) ||
        periodStart.getMonth() !== Number(q[1]) - 1 ||
        periodStart.getDate() !== Number(q[2])) {
      throw new Error('Bill period start is not a valid calendar date.');
    }
  }
  var periodEnd = addDays_(addMonths_(periodStart, months), -1);
  var invoiceNo = nextInvoiceNo_(gstVals, today);
  var cgst = round2_(amount * 0.09);
  var sgst = round2_(amount * 0.09);
  var total = round2_(amount + cgst + sgst);
  var invoiceDateDoc = fmtShort_(today).replace(/\//g, '-');
  var seed = {
    buyer: name,
    address: address,
    gstin: gstin,
    userid: userid,
    invoiceNo: invoiceNo,
    invoiceDate: invoiceDateDoc,
    periodStart: fmtDMY_(periodStart),
    periodEnd: fmtDMY_(periodEnd),
    desc: desc,
    amount: amount,
    cgst: cgst,
    sgst: sgst,
    total: total,
    mode: mode
  };

  var copyId = '';
  var archive = null;
  var mapSaved = false;
  try {
    // Copy the blank master template, never another customer's live invoice.
    copyId = templateFile.makeCopy(docName, docsRoot).getId();

    var baselineHash = null;
    try { baselineHash = blobHash_(exportPdf_(copyId)); } catch (baselineError) { /* optional freshness baseline */ }

    populateNewCustomerTemplate_(copyId, seed);
    verifyNewCustomerTemplate_(copyId, seed);

    // Archive into the same customer key used by both sheets and DOCS_MAP.
    archive = exportAndArchive_(copyId, sheetName, invoiceNo,
      String(fmtShort_(today)).replace(/\//g, '-'), baselineHash);
    var pdfBase64 = Utilities.base64Encode(archive.blob.getBytes());

    // Persist records only after the document and PDF have passed verification.
    var firstGstRow = appendGstRecord_(gst, gCols, {
      invoiceNo: invoiceNo,
      invoiceDate: fmtShort_(today),
      customer: sheetName,
      gstin: gstin,
      desc: desc,
      amount: amount,
      cgst: cgst,
      sgst: sgst,
      total: total,
      mode: mode,
      payment_status: 'Paid'
    });
    writeBillPeriod_(gst, ensureBillPeriodCol_(gst), firstGstRow - 1, periodStart, periodEnd);

    var newRow = up.getLastRow() + 1;
    up.getRange(newRow, iSrNo + 1).setValue(nextUpcomingSrNo_(up, iSrNo));
    up.getRange(newRow, iName + 1).setValue(sheetName);
    var expCell = up.getRange(newRow, iExp + 1);
    expCell.setNumberFormat('dd/mm/yy');
    expCell.setValue(periodEnd);
    if (iDur >= 0) up.getRange(newRow, iDur + 1).setValue(durationText_(months));
    if (address && iAddr >= 0) up.getRange(newRow, iAddr + 1).setValue(address);
    if (userid && iUser >= 0) up.getRange(newRow, iUser + 1).setValue(userid);
    sortUpcomingByExpire_(up, iExp);

    map[sheetName] = copyId;
    saveDocsMap_(map);
    mapSaved = true;

    return {
      ok: true,
      customer: sheetName,
      doc_name: docName,
      invoice_no: invoiceNo,
      invoice_date: fmtShort_(today),
      period: { start: fmtDMY_(periodStart), end: fmtDMY_(periodEnd) },
      new_expire: fmtDMY_(periodEnd),
      amount: amount,
      cgst: cgst,
      sgst: sgst,
      total: total,
      desc: desc,
      mode: mode,
      duration_months: months,
      doc_id: copyId,
      pdf_source: 'master-template (verified)',
      pdf_file: archive.file,
      pdf_folder: archive.folderUrl,
      pdf_name: sanitizeFile_(sheetName) + '_BILL.pdf',
      pdf_base64: pdfBase64
    };
  } catch (err) {
    // Best-effort rollback of any partially persisted state and created Drive assets.
    if (mapSaved) {
      try { delete map[sheetName]; saveDocsMap_(map); } catch (mapRollbackError) { Logger.log('Onboarding map rollback failed: ' + mapRollbackError); }
    }
    try { rollbackOnboardingRows_(gst, up, invoiceNo, sheetName); }
    catch (rowRollbackError) { Logger.log('Onboarding sheet rollback failed: ' + rowRollbackError); }
    if (archive && archive.fileId) {
      try { DriveApp.getFileById(archive.fileId).setTrashed(true); } catch (pdfRollbackError) { Logger.log('Onboarding PDF cleanup failed: ' + pdfRollbackError); }
    }
    if (copyId) {
      try { DriveApp.getFileById(copyId).setTrashed(true); } catch (docRollbackError) { Logger.log('Onboarding Doc cleanup failed: ' + docRollbackError); }
    }
    throw new Error('Customer onboarding did not complete. Any created bill, PDF and sheet rows were rolled back where possible. ' +
      String(err && err.message || err));
  }
}


/**
 * Fills all required tokens in a fresh copy of the master invoice template.
 */
function populateNewCustomerTemplate_(docId, bill) {
  var descriptionLines = String(bill.desc || '').split(/\r?\n/)
    .map(function (line) { return line.replace(/\s+/g, ' ').trim(); })
    .filter(function (line) { return !!line; });
  var firstPlanLine = descriptionLines.length ? descriptionLines[0] : '';
  var secondPlanLine = descriptionLines.length > 1 ? descriptionLines.slice(1).join(' ') : '';
  var replacements = {
    BUYER_NAME: bill.buyer,
    BUYER_ADDRESS: bill.address,
    BUYER_GSTIN_LINE: bill.gstin ? 'GSTIN / UIN : - ' + bill.gstin : '',
    BUYER_STATE: 'Maharashtra',
    BUYER_STATE_CODE: '27',
    INVOICE_DATE: bill.invoiceDate,
    INVOICE_NO: bill.invoiceNo,
    USER_ID: bill.userid,
    PLACE_OF_SUPPLY: 'MAHARASHTRA',
    PLACE_STATE_CODE: '27',
    BILL_PERIOD: bill.periodStart + ' to ' + bill.periodEnd,
    PLAN_LINE_1: firstPlanLine,
    PLAN_LINE_2: secondPlanLine,
    CHARGES: money_(bill.amount),
    LINE_AMOUNT: money_(bill.amount),
    BASIC_AMOUNT: money_(bill.amount),
    CGST_AMOUNT: money_(bill.cgst),
    SGST_AMOUNT: money_(bill.sgst),
    ROUND_OFF: money_(0),
    TOTAL_AMOUNT: money_(bill.total),
    AMOUNT_IN_WORDS: rupeeWords_(bill.total) + ' RUPEES ONLY/_'
  };

  var before = docText_(docId);
  var missingTokens = [];
  Object.keys(replacements).forEach(function (key) {
    if (before.indexOf('[[' + key + ']]') < 0) missingTokens.push('[[' + key + ']]');
  });
  if (missingTokens.length) {
    throw new Error('The configured master template is missing required placeholders: ' + missingTokens.join(', '));
  }

  Object.keys(replacements).forEach(function (key) {
    replaceTextInDoc_(docId, '[[' + key + ']]', String(replacements[key]));
  });

  var after = docText_(docId);
  var leftovers = after.match(/\[\[[A-Z0-9_]+\]\]/g) || [];
  if (leftovers.length) {
    throw new Error('Template placeholders were not fully replaced: ' + leftovers.join(', '));
  }
}


/**
 * Verifies the actual populated Google Doc before its PDF is exported.
 */
function verifyNewCustomerTemplate_(docId, bill) {
  var text = docText_(docId);
  function normalized(value) {
    return String(value === null || value === undefined ? '' : value)
      .replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  }
  function requireText(label, value) {
    var expected = normalized(value);
    if (expected && normalized(text).indexOf(expected) < 0) {
      throw new Error('Template verification failed for ' + label + '. Expected value was not found in the Google Doc.');
    }
  }

  var printedBuyer = buyerNameInDoc_(docId);
  if (normalized(printedBuyer) !== normalized(bill.buyer)) {
    throw new Error('Template verification failed for Buyer name. Expected "' + bill.buyer + '", found "' + (printedBuyer || '(blank)') + '".');
  }
  var printedUserId = userIdInDoc_(docId);
  if (normalized(printedUserId) !== normalized(bill.userid)) {
    throw new Error('Template verification failed for User ID. Expected "' + bill.userid + '", found "' + (printedUserId || '(blank)') + '".');
  }
  var printedDate = invoiceDateTextInDoc_(docId);
  if (normalized(printedDate) !== normalized(bill.invoiceDate)) {
    throw new Error('Template verification failed for Invoice Date. Expected "' + bill.invoiceDate + '", found "' + (printedDate || '(blank)') + '".');
  }

  requireText('Buyer address', bill.address);
  if (bill.gstin) requireText('Buyer GSTIN / UIN', bill.gstin);
  requireText('Invoice number', bill.invoiceNo);
  requireText('Bill period', bill.periodStart + ' to ' + bill.periodEnd);
  requireText('Plan description line 1', String(bill.desc).split(/\r?\n/).filter(Boolean)[0] || '');
  var descLines = String(bill.desc).split(/\r?\n/).map(function (line) { return line.trim(); }).filter(Boolean);
  if (descLines.length > 1) requireText('Plan description', descLines.slice(1).join(' '));
  requireText('Basic amount', money_(bill.amount));
  requireText('CGST amount', money_(bill.cgst));
  requireText('SGST amount', money_(bill.sgst));
  requireText('Invoice total', money_(bill.total));
  requireText('Amount in words', rupeeWords_(bill.total) + ' RUPEES ONLY/_');
  requireText('State name', 'Maharashtra');
  requireText('State code', '27');
  requireText('Place of supply', 'MAHARASHTRA');

  return { ok: true, verified: true };
}


/**
 * Remove only rows belonging to this newly reserved invoice/customer.
 * Called only while the script lock is held.
 */
function rollbackOnboardingRows_(gst, up, invoiceNo, customer) {
  if (gst && gst.getLastRow() > 1) {
    var gv = gst.getDataRange().getValues();
    var gi = colOf_(gv[0], 'Invoice No');
    var gc = colOf_(gv[0], 'Customer Name');
    if (gi >= 0 && gc >= 0) {
      for (var r = gv.length - 1; r >= 1; r--) {
        if (String(gv[r][gi] || '').trim() === String(invoiceNo) &&
            String(gv[r][gc] || '').trim().toLowerCase() === String(customer).toLowerCase()) {
          gst.deleteRow(r + 1);
        }
      }
    }
  }
  if (up && up.getLastRow() > 1) {
    var uv = up.getDataRange().getValues();
    var ui = colOf_(uv[0], 'Customer name');
    if (ui >= 0) {
      for (var ur = uv.length - 1; ur >= 1; ur--) {
        if (String(uv[ur][ui] || '').trim().toLowerCase() === String(customer).toLowerCase()) {
          up.deleteRow(ur + 1);
        }
      }
    }
  }
}

/* --------------------------------------------------- doc health checks ---- */

/** check_docs: for every customer, verify the mapped doc contains exactly the
 *  values the last GST record says it should (read-only). */
function checkDocs_(body) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var up = ss.getSheetByName(UPCOMING_SHEET);
  var gst = ss.getSheetByName(GST_SHEET);
  var upVals = up.getDataRange().getValues();
  var iName = colOf_(upVals[0], 'Customer name'), iExp = colOf_(upVals[0], 'Expire Date');
  var iDur = colOf_(upVals[0], COL_DURATION);
  var gstVals = gst.getDataRange().getValues();
  var raw = PROPS.getProperty('DOCS_MAP') || '{}';
  var map; try { map = JSON.parse(raw); } catch (e) { map = {}; }
  var exportTest = body.export_test === true;
  var out = [];
  for (var r = 1; r < upVals.length; r++) {
    var cust = String(upVals[r][iName] || '').trim();
    if (!cust) continue;
    var docId = map[cust] || '';
    var entry = { customer: cust, doc_id: docId || '(not mapped)' };
    if (!docId) { entry.status = 'NOT MAPPED'; out.push(entry); continue; }
    try {
      var text = docText_(docId);
      var last = findLastGstRecord_(gstVals, cust);
      var expire = asDate_(upVals[r][iExp]);
      var months = parseDuration_(iDur >= 0 ? String(upVals[r][iDur] || '') : '');
      var end = expire, start = (end && months > 0) ? addDays_(addMonths_(end, -months), 1) : null;
      var checks = {
        'Invoice No':  last ? last.invoiceNo : '',
        'Invoice Date': last ? String(last.invoiceDate || '').replace(/\//g, '-') : '',
        'Bill Period': start ? (fmtDMY_(start) + ' to ' + fmtDMY_(end)) : '',
        'Amount':      last ? money_(last.amount) : '',
        'CGST/SGST':   last ? money_(last.cgst) : '',
        'Total':       last ? money_(last.total) : '',
        'Words':       last ? (rupeeWords_(last.total) + ' RUPEES ONLY/_') : ''
      };
      var missing = [];
      for (var k in checks) {
        if (!checks[k]) continue;
        if (text.indexOf(checks[k]) >= 0) continue;
        // Bill Period: accept any period line whose END date matches the
        // expiry (start-date conventions can differ by a day, e.g. a
        // 01/11->30/11 month can back-compute as 31/10->30/11)
        if (k === 'Bill Period' && expire) {
          var pm = text.match(new RegExp('\\d{2}/\\d{2}/\\d{4} to ' +
            escRegex_(fmtDMY_(expire))));
          if (pm) continue;
        }
        missing.push(k + ' ("' + String(checks[k]).slice(0, 30) + '")');
      }
      entry.lastInvoice = last ? last.invoiceNo : '(none in GST RECORDS)';
      if (exportTest) {
        try { entry.export_test = exportPdf_(docId).getBytes().length + ' bytes'; }
        catch (e2) { entry.export_test = 'FAILED: ' + String(e2 && e2.message || e2).slice(0, 90); }
      }
      entry.status = missing.length ? 'NEEDS ATTENTION — missing: ' + missing.join('; ') : 'OK';
    } catch (e) {
      entry.status = 'ERROR: ' + String(e && e.message || e).slice(0, 140);
    }
    out.push(entry);
  }
  return { ok: true, results: out };
}

/** update_payment: set Payment Status / Payment Date on ONE GST RECORDS row.
 *  body: { customer, invoice_no, payment_status: 'Paid'|'Pending',
 *          payment_date: 'yyyy-mm-dd' (used when status = Paid; defaults today) }
 *  The row is located by invoice no + best fuzzy customer match (invoice nos
 *  are NOT unique across customers — e.g. two Skin Cities share ST/048). */
function updatePayment_(body) {
  var cust = String(body.customer || '').trim();
  var inv = String(body.invoice_no || '').trim();
  var status = String(body.payment_status || '').trim();
  if (!cust || !inv) throw new Error('customer and invoice_no are required');
  if (status !== 'Paid' && status !== 'Pending')
    throw new Error("payment_status must be 'Paid' or 'Pending'");
  var payDate = null;
  if (status === 'Paid') {
    var pdRaw = String(body.payment_date || '').trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(pdRaw)) {
      var p = pdRaw.split('-');
      payDate = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    } else {
      payDate = startOfDay_(new Date());
    }
  }
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var gst = ss.getSheetByName(GST_SHEET);
  var vals = gst.getDataRange().getValues();
  var iStatus = colOf_(vals[0], 'Payment Status');
  var iPayDate = colOf_(vals[0], 'Payment Date');
  if (iStatus < 0) throw new Error('GST RECORDS has no "Payment Status" column');
  var row = -1, bestScore = 0;
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][0] || '').trim() !== inv) continue;
    var s = nameScore_(cust, String(vals[r][2] || ''));
    if (s > bestScore) { bestScore = s; row = r; }
  }
  if (row < 0) throw new Error('no GST RECORDS row with invoice ' + inv + ' for ' + cust);
  var prev = String(vals[row][iStatus] || '').trim();
  gst.getRange(row + 1, iStatus + 1).setValue(status);
  var dateStr = '';
  if (iPayDate >= 0) {
    var cellD = gst.getRange(row + 1, iPayDate + 1);
    if (status === 'Paid' && payDate) {
      cellD.setNumberFormat('dd/mm/yy');
      cellD.setValue(payDate);
      dateStr = fmtDMY_(payDate);
    } else if (status === 'Pending') {
      cellD.setValue('');                    // unpaid -> no payment date
    }
  }
  return { ok: true, customer: cust, invoice_no: inv,
           previous_status: prev, payment_status: status,
           payment_date: dateStr, row: row + 1 };
}

/** repair_doc: bring a customer's Google Doc back in sync with their LATEST
 *  GST RECORDS entry. Needed for docs that missed edits during the era when
 *  server-side doc writes failed: the doc still shows an OLDER invoice, so a
 *  normal bill would fail (the engine looks for the last invoice's values).
 *  Method (all read-only recon first):
 *    1. read the doc, take its actual invoice no (ST/xxx/…)
 *    2. find THAT invoice's GST record -> the doc's true current values
 *    3. read the doc's actual period line ("dd/mm/yyyy to dd/mm/yyyy")
 *    4. build old->new pairs (doc state -> latest record) and apply through
 *       the normal verified engine (aborts, nothing saved, on any mismatch).
 *  body: { customer: "..." } or { all: true } for every stale mapped customer. */
function repairDoc_(body) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var up = ss.getSheetByName(UPCOMING_SHEET);
  var gst = ss.getSheetByName(GST_SHEET);
  var upVals = up.getDataRange().getValues();
  var gstVals = gst.getDataRange().getValues();
  var iName = colOf_(upVals[0], 'Customer name'), iExp = colOf_(upVals[0], 'Expire Date');
  var iDur = colOf_(upVals[0], COL_DURATION);
  var map; try { map = JSON.parse(PROPS.getProperty('DOCS_MAP') || '{}'); } catch (e) { map = {}; }
  var only = String(body.customer || '').trim();
  var out = [];
  for (var r = 1; r < upVals.length; r++) {
    var cust = String(upVals[r][iName] || '').trim();
    if (!cust) continue;
    if (only && cust.toLowerCase() !== only.toLowerCase()) continue;
    var docId = map[cust] || '';
    if (!docId) { if (only) out.push({ customer: cust, error: 'not mapped' }); continue; }

    var entry = { customer: cust };
    try {
      var text = docText_(docId);
      var invM = text.match(/ST\/\d{1,4}\/\d{4}-\d{1,2}/);
      var last = findLastGstRecord_(gstVals, cust);
      if (!invM || !last) throw new Error('cannot read invoice no from doc or no GST record');
      entry.doc_invoice = invM[0];
      entry.latest_invoice = last.invoiceNo;

      // the invoice the doc actually shows -> its GST row = the doc's true state
      var old = null;
      if (invM[0] === last.invoiceNo) {
        old = { invoiceNo: last.invoiceNo, invoiceDate: String(last.invoiceDate || ''),
                customer: last.customer, desc: last.desc, amount: last.amount,
                cgst: last.cgst, sgst: last.sgst, total: last.total };
      } else {
        for (var g = 1; g < gstVals.length; g++) {
          if (String(gstVals[g][0] || '').trim() === invM[0]) {
            old = { invoiceNo: invM[0],
                    invoiceDate: String(gstVals[g][1] || '').trim(),
                    customer: String(gstVals[g][2] || '').trim(),
                    desc: String(gstVals[g][4] || '').trim(),
                    amount: num_(gstVals[g][5]), cgst: num_(gstVals[g][6]),
                    sgst: num_(gstVals[g][7]), total: num_(gstVals[g][8]) };
            break;
          }
        }
        if (!old || !(old.total > 0)) throw new Error('doc shows invoice ' + invM[0] +
          ' which has no GST RECORDS row — cannot derive its current values');
      }

      // the doc's ACTUAL period + words lines (their formatting may drift from
      // what the engine computes) — used as the replacement source so the edit
      // always finds its anchor
      var perM = text.match(/\d{2}\/\d{2}\/\d{4} to \d{2}\/\d{2}\/\d{4}/);
      var wordM = text.match(/[A-Z][A-Z ,\-]*RUPEES ONLY[^\n]*/);
      var expire = asDate_(upVals[r][iExp]);
      var months = parseDuration_(iDur >= 0 ? String(upVals[r][iDur] || '') : '');
      if (!(months > 0)) throw new Error('no parseable Duration in UPCOMING BILLS');
      var newStart = expire ? addDays_(addMonths_(expire, -months), 1) : null;
      if (!newStart) throw new Error('no parseable Expire Date in UPCOMING BILLS');
      var bill = { invoiceNo: last.invoiceNo,
                   invoiceDate: String(last.invoiceDate || ''),
                   periodStart: fmtDMY_(newStart), periodEnd: fmtDMY_(expire),
                   amount: last.amount, cgst: last.cgst, sgst: last.sgst,
                   total: last.total, desc: last.desc };
      var pairs = docEditPairs_(old, null, 0, bill);
      pairs.forEach(function (p) {
        if (p.name === 'Bill Period' && perM) p.old = perM[0];
        if (p.name === 'Invoice Date' && old.invoiceDate) p.old = String(old.invoiceDate).replace(/\//g, '-');
        if (p.name === 'Words' && wordM) p.old = wordM[0];
      });
      var edit = editCustomerDoc_(docId, pairs);
      entry.status = invM[0] === last.invoiceNo
        ? (edit.edited && edit.edited.length ? 'polished: ' + edit.edited.join(', ') : 'already current')
        : 'repaired ' + invM[0] + ' -> ' + last.invoiceNo;
      entry.edited = edit.edited; entry.skipped = edit.skipped;
    } catch (e) {
      entry.status = 'ERROR: ' + String(e && e.message || e).slice(0, 160);
    }
    out.push(entry);
    if (only) break;
  }
  if (only && !out.length) throw new Error('customer not found in UPCOMING BILLS: ' + only);
  return { ok: true, results: out };
}

/** rebuild_doc: recreate one customer's doc from the FIXED template
 *  (GST_BILL_TEMPLATE.docx converted to a Google Doc), pre-filled with that
 *  customer's LAST GST record, and map it. Use when a doc is broken/unusable. */
function rebuildDoc_(body) {
  var name = String(body.customer || '').trim();
  if (!name) throw new Error('customer is required');
  var tmplId = PROPS.getProperty('BILL_TEMPLATE_DOC_ID');
  if (!tmplId) throw new Error('BILL_TEMPLATE_DOC_ID not set — upload the FIXED ' +
    'GST_BILL_TEMPLATE.docx to the telecom Drive, open it as a Google Doc and register it first');
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var up = ss.getSheetByName(UPCOMING_SHEET);
  var gst = ss.getSheetByName(GST_SHEET);
  var upVals = up.getDataRange().getValues();
  var iName = colOf_(upVals[0], 'Customer name'), iExp = colOf_(upVals[0], 'Expire Date');
  var iDur = colOf_(upVals[0], COL_DURATION), iAddr = colOf_(upVals[0], COL_ADDRESS);
  var iUser = colOf_(upVals[0], COL_USERID);
  var row = -1;
  for (var r = 1; r < upVals.length; r++) {
    if (String(upVals[r][iName] || '').trim().toLowerCase() === name.toLowerCase()) { row = r; break; }
  }
  if (row < 0) throw new Error('customer not found in UPCOMING BILLS: ' + name);
  var custName = String(upVals[row][iName] || '').trim();
  var expire = asDate_(upVals[row][iExp]);
  var last = findLastGstRecord_(gst.getDataRange().getValues(), custName);
  if (!last) throw new Error('no GST RECORDS entry for ' + custName + ' — cannot rebuild');
  var months = parseDuration_(iDur >= 0 ? String(upVals[row][iDur] || '') : '');
  var end = expire, start = (end && months > 0) ? addDays_(addMonths_(end, -months), 1) : null;

  var bill = {
    buyer: custName,
    address: iAddr >= 0 ? String(upVals[row][iAddr] || '').trim() : '',
    gstin: last.gstin,
    userid: iUser >= 0 ? String(upVals[row][iUser] || '').trim() : '',
    invoiceNo: last.invoiceNo,
    invoiceDate: String(last.invoiceDate || '').replace(/\//g, '-'),
    periodStart: start ? fmtDMY_(start) : '-',
    periodEnd: end ? fmtDMY_(end) : '-',
    desc: last.desc,
    amount: last.amount, cgst: last.cgst, total: last.total
  };
  var vals = templateValues_(bill);
  var toks = Object.keys(vals);
  var copyId = DriveApp.getFileById(tmplId).makeCopy(sanitizeFile_(custName),
                                                     docsFolder_()).getId();
  try {
    Utilities.sleep(1500);                    // let the fresh copy settle
    var doc = null, openErr = null;
    for (var att = 0; att < 3 && !doc; att++) {
      try { doc = DocumentApp.openById(copyId); }
      catch (e) { openErr = e; Utilities.sleep(2500); }
    }
    if (!doc) throw new Error('cannot open the new copy: ' +
      String(openErr && openErr.message || openErr));
    var b = doc.getBody();
    for (var tok in vals) {
      b.replaceText(escRegex_(tok), String(vals[tok]).replace(/\$/g, '$$$$'));
    }
    doc.saveAndClose();
    var text = docText_(copyId);
    var missing = toks.filter(function (tok) { return text.indexOf(tok) >= 0; });
    if (missing.length) throw new Error('tokens not filled: ' + missing.join(', '));
  } catch (err) {
    try { DriveApp.getFileById(copyId).setTrashed(true); } catch (e) { /* keep going */ }
    throw err;
  }
  var raw = PROPS.getProperty('DOCS_MAP') || '{}';
  var map; try { map = JSON.parse(raw); } catch (e) { map = {}; }
  map[custName] = copyId;
  saveDocsMap_(map);
  var pdfBytes = 0;
  try { pdfBytes = exportPdf_(copyId).getBytes().length; } catch (e) {}
  return { ok: true, customer: custName, doc_id: copyId,
           doc_name: DriveApp.getFileById(copyId).getName(),
           export_test: pdfBytes + ' bytes', mapped: true,
           note: 'if an older doc with the same name exists in the folder, delete it to avoid confusion' };
}

/* --------------------------------------------------- archived PDF history - */

/** list_pdfs: every archived PDF of one customer (newest first). */
function listPdfs_(body) {
  var cust = String(body.customer || '').trim();
  if (!cust) throw new Error('customer is required');
  var folder = customerFolder_(cust, false);
  if (!folder) return { ok: true, customer: cust, files: [], folder_url: '' };
  var files = [];
  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    files.push({ name: f.getName(), id: f.getId(),
                 size: f.getSize(), created: f.getDateCreated().getTime() });
  }
  files.sort(function (a, b) { return b.created - a.created; });
  return { ok: true, customer: cust, folder_url: folder.getUrl(), files: files };
}

/** get_pdf: fetch one archived PDF by invoice number (history viewer). */
function getPdf_(body) {
  var invNo = String(body.invoice_no || '').trim();
  if (!invNo) throw new Error('invoice_no is required');
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var gst = ss.getSheetByName(GST_SHEET);
  var vals = gst.getDataRange().getValues();
  var gRow = -1;
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][0] || '').trim() === invNo) { gRow = r; break; }
  }
  if (gRow < 0) throw new Error('invoice not found in GST RECORDS: ' + invNo);
  var cust = String(vals[gRow][colOf_(vals[0], 'Customer Name')] || '').trim();
  var folder = customerFolder_(cust, false);
  if (!folder) throw new Error('no archived PDFs folder for ' + cust);
  var prefix = sanitizeFile_(invNo);
  var it = folder.getFiles(), found = null, newestCreated = -1;
  while (it.hasNext()) {
    var f = it.next();
    if (String(f.getName()).indexOf(prefix) === 0) {
      var created = f.getDateCreated().getTime();
      if (created >= newestCreated) { found = f; newestCreated = created; }
    }
  }
  if (!found) throw new Error('no archived PDF for ' + invNo +
    ' in BILLS/' + sanitizeFile_(cust) +
    ' — PDFs are archived to Drive from now on; older ones exist only on the PC');
  return { ok: true, invoice_no: invNo, customer: cust,
           pdf_name: found.getName(),
           pdf_base64: Utilities.base64Encode(found.getBlob().getBytes()) };
}

/* -------------------------------------------------------------------- make_pdf */

/**
 * Returns the ARCHIVED PDF for an invoice (from BILLS/<Customer>/ on Drive).
 * No sheet writes, no regeneration — the customer docs always hold the
 * LATEST bill, so old invoices come from the archive.
 */
function makePdf_(body) {
  var invNo = String(body.invoice_no || '').trim();
  if (!invNo) throw new Error('invoice_no is required');
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var gst = ss.getSheetByName(GST_SHEET);
  var vals = gst.getDataRange().getValues();
  var gRow = -1;
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][0] || '').trim() === invNo) { gRow = r; break; }
  }
  if (gRow < 0) throw new Error('invoice not found in GST RECORDS: ' + invNo);
  var cust = String(vals[gRow][colOf_(vals[0], 'Customer Name')] || '').trim();
  var folder = customerFolder_(cust, false);
  if (!folder) throw new Error('no archived PDFs folder for ' + cust);
  var prefix = sanitizeFile_(invNo);
  var it = folder.getFiles(), found = null;
  while (it.hasNext()) {
    var f = it.next();
    if (String(f.getName()).indexOf(prefix) === 0) { found = f; break; }
  }
  if (!found) throw new Error('no archived PDF for ' + invNo +
    ' in BILLS/' + sanitizeFile_(cust) +
    ' — bills are archived from now on; older PDFs exist only on the PC');
  return { ok: true, invoice_no: invNo, customer: cust,
           pdf_name: found.getName(),
           pdf_base64: Utilities.base64Encode(found.getBlob().getBytes()) };
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
  step('copy export fallback', function () {
    if (!copyId) throw new Error('no copy made');
    return exportPdf_(copyId).getBytes().length + ' bytes';
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
  step('ORIGINAL export fallback', function () {
    if (!docId) throw new Error('no docId set');
    return exportPdf_(docId).getBytes().length + ' bytes';
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

/** debug_rebuild: runs the rebuild_doc pipeline STEP BY STEP with timing and
 *  reports exactly which call fails. Creates + trashes one template copy.
 *  Params (for bisection — hard aborts lose the response, so vary per request):
 *    customer   — which customer's data to use (default Ace Apparel)
 *    upto       — replace only the FIRST N tokens (default: all)
 *    skip_export — true: don't run the export test
 */
function debugRebuild_(body) {
  var steps = [];
  function step(name, fn) {
    var t0 = Date.now();
    try {
      var v = fn();
      steps.push({ step: name, ok: true, ms: Date.now() - t0,
                   info: v === undefined ? '' : String(v) });
    } catch (e) {
      steps.push({ step: name, ok: false, ms: Date.now() - t0,
                   error: String(e && e.message || e).slice(0, 200) });
    }
  }
  // --- plain-doc write test: does ANY Google Doc accept writes here? -------
  var name = String(body.customer || 'Ace Apparel').trim();
  var upto = (body.upto === undefined || body.upto === null) ? 9999 : Number(body.upto);
  var skipExport = body.skip_export === true;
  var tmplId = PROPS.getProperty('BILL_TEMPLATE_DOC_ID') || '(not set)';
  var folderId = PROPS.getProperty('DOCS_FOLDER_ID') || '(not set)';
  var tmplInfo = '';
  if (body.plain_test) {
    step('plain doc write test', function () {
      var d = DocumentApp.create('debug-plain-' + Date.now());
      var id = d.getId();
      try {
        d.getBody().setText('hello {{NAME}} test');
        d.getBody().replaceText(escRegex_('{{NAME}}'), 'WORLD');
        var t = d.getBody().getText();
        if (t.indexOf('WORLD') < 0) throw new Error('replace did not apply');
        return 'plain doc write OK: "' + t + '"';
      } finally {
        try { DriveApp.getFileById(id).setTrashed(true); } catch (e2) {}
      }
    });
  }
  // --- copy-blocks test: template READ-ONLY -> copy body blocks into a fresh
  //     doc -> WRITE in the fresh doc. If this survives, rebuild = copy+fill. --
  if (body.mode === 'copyblocks') {
    if (body.mode === 'copyblocks') {
      step('copy blocks to fresh doc', function () {
        var tmpl = DocumentApp.openById(tmplId);            // read-only access
        var src = tmpl.getBody();
        var fresh = DocumentApp.create('debug-copyblocks-' + Date.now());
        var fid = fresh.getId();
        try {
          var dst = fresh.getBody();
          var n = 0, imgs = 0;
          for (var i = 0; i < src.getNumChildren(); i++) {
            var el = src.getChild(i).copy();
            var t = el.getType();
            if (t === DocumentApp.ElementType.PARAGRAPH) {
              if (el.getNumChildren && el.getNumChildren() > 0) {
                for (var j = 0; j < el.getNumChildren(); j++) {
                  if (el.getChild(j).getType() === DocumentApp.ElementType.INLINE_IMAGE) imgs++;
                }
              }
              dst.appendParagraph(el.asParagraph());
            } else if (t === DocumentApp.ElementType.TABLE) {
              dst.appendTable(el.asTable());
            } else if (t === DocumentApp.ElementType.LIST_ITEM) {
              dst.appendListItem(el.asListItem());
            } else if (t === DocumentApp.ElementType.HORIZONTAL_RULE) {
              dst.appendHorizontalRule(el.asHorizontalRule());
            } else {
              dst.appendParagraph('');
            }
            n++;
          }
          var before = dst.getText();
          dst.replaceText(escRegex_('{{BUYER}}'), 'ACE APPAREL WRITE-TEST');
          var after = dst.getText();
          if (after.indexOf('ACE APPAREL WRITE-TEST') < 0) throw new Error('write did not apply');
          var pdf = exportPdf_(fid).getBytes().length;
          return 'OK — copied ' + n + ' blocks (' + imgs + ' images), write applied, export ' + pdf + ' bytes';
        } finally {
          try { DriveApp.getFileById(fid).setTrashed(true); } catch (e2) {}
        }
      });
    }
    return { ok: true, customer: name, template_info: tmplInfo, debug: steps };
  }
  step('props', function () { return 'template: ' + tmplId + ', folder: ' + folderId +
    ', engine: DocumentApp (pure Apps Script)'; });
  step('sheet', function () { return SpreadsheetApp.openById(SHEET_ID).getName(); });
  var cust = '';
  var last = null;
  var expire = null, months = 0;
  step('find customer', function () {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var up = ss.getSheetByName(UPCOMING_SHEET);
    var vals = up.getDataRange().getValues();
    var iName = colOf_(vals[0], 'Customer name');
    for (var r = 1; r < vals.length; r++) {
      if (String(vals[r][iName] || '').trim().toLowerCase() === name.toLowerCase()) {
        cust = String(vals[r][iName] || '').trim();
        expire = asDate_(vals[r][colOf_(vals[0], 'Expire Date')]);
        months = parseDuration_(String(vals[r][colOf_(vals[0], COL_DURATION)] || ''));
        return 'row ' + r + ': ' + cust + ', expire ' + expire + ', months ' + months;
      }
    }
    throw new Error('not found: ' + name);
  });
  step('last gst record', function () {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var gst = ss.getSheetByName(GST_SHEET);
    last = findLastGstRecord_(gst.getDataRange().getValues(), cust);
    if (!last) throw new Error('no GST record for ' + cust);
    return last.invoiceNo + ' @ ' + last.invoiceDate + ', total ' + money_(last.total);
  });
  var tmplInfo = '';
  step('template open', function () {
    var d = DocumentApp.openById(tmplId);
    var h = d.getHeader();
    tmplInfo = 'paras ' + d.getBody().getParagraphs().length +
      ', tables ' + d.getBody().getTables().length +
      ', header image elements ' + (h ? h.getImages().length : 'no header') +
      ', body text starts: "' + String(d.getBody().getText() || '').slice(0, 90) + '"';
    return tmplInfo;
  });
  var copyId = '';
  step('makeCopy into docs folder', function () {
    var c = DriveApp.getFileById(tmplId).makeCopy(sanitizeFile_(name),
                                                  DriveApp.getFolderById(folderId));
    copyId = c.getId();
    return copyId;
  });
  step('openById(copy) read probe', function () {
    var d = DocumentApp.openById(copyId);          // READ only — proven safe
    return 'opened, paras ' + d.getBody().getParagraphs().length;
  });
  step('replaceText tokens', function () {
    var vals = templateValues_({ buyer: cust, gstin: last.gstin,
      invoiceNo: last.invoiceNo, invoiceDate: String(last.invoiceDate || '').replace(/\//g, '-'),
      periodStart: '-', periodEnd: '-', desc: last.desc,
      amount: last.amount, cgst: last.cgst, total: last.total });
    var keys = Object.keys(vals);
    var t0 = Date.now();
    var done = [];
    var d2 = DocumentApp.openById(copyId);
    var b = d2.getBody();
    for (var i = 0; i < keys.length && i < upto; i++) {
      b.replaceText(escRegex_(keys[i]), String(vals[keys[i]]).replace(/\$/g, '$$$$'));
      done.push(keys[i]);
    }
    d2.saveAndClose();
    return (Date.now() - t0) + ' ms — replaced ' + done.length + '/' + keys.length +
      ' tokens (' + done.join(', ') + ')';
  });
  step('token check', function () {
    var text = docText_(copyId);                   // DocumentApp read-back (post-edit)
    var left = [];
    for (var k in templateValues_({ buyer: '', gstin: '', invoiceNo: '', invoiceDate: '',
      periodStart: '', periodEnd: '', desc: '', amount: 0, cgst: 0, total: 0 })) {
      if (text.indexOf(k) >= 0) left.push(k);
    }
    if (left.length) throw new Error('unfilled tokens: ' + left.join(', '));
    return 'all filled';
  });
  if (skipExport) {
    step('export test', function () { return 'skipped'; });
  } else {
    step('export test', function () {
      return exportPdf_(copyId).getBytes().length + ' bytes';
    });
  }
  step('trash copy', function () {
    if (copyId) DriveApp.getFileById(copyId).setTrashed(true);
  });
  return { ok: true, customer: name, template_info: tmplInfo, debug: steps };
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

/** Convert full document wording into the compact label stored in GST
 *  RECORDS, e.g. "100 Mbps (1 Month)". Empty means it could not be parsed. */
function conciseServiceDescription_(description) {
  var text = String(description || '')
    .replace(/&nbsp;|&#xA0;|&#160;/gi, ' ')
    .replace(/[\u00A0\u202F]/g, ' ')
    .replace(/\s+/g, ' ').trim();
  if (!text) return '';
  var speed = text.match(/\b(\d+(?:\.\d+)?)\s*(gbps|gb\/s|mbps|mb\/s|kbps|kb\/s)\b/i);
  var months = parseDuration_(text);
  if (!speed || !(months > 0)) return '';
  var unit = /^gb/i.test(speed[2]) ? 'Gbps' :
             /^kb/i.test(speed[2]) ? 'Kbps' : 'Mbps';
  var duration = months === 1 ? '1 Month' : months + ' Months';
  return speed[1] + ' ' + unit + ' (' + duration + ')';
}

/** Latest prior description for this customer, excluding the edited invoice.
 *  Used only when the current doc wording cannot be normalized. */
function previousServiceDescription_(gstVals, customer, excludeInvoice) {
  if (!gstVals || !gstVals.length) return '';
  var iInv = colOf_(gstVals[0], 'Invoice No');
  var iName = colOf_(gstVals[0], 'Customer Name');
  var iDesc = colOf_(gstVals[0], 'Service Description');
  var iDate = colOf_(gstVals[0], 'Invoice Date');
  if (iInv < 0 || iName < 0 || iDesc < 0) return '';
  var best = '', bestScore = 0, bestDate = null, bestRow = -1;
  for (var r = 1; r < gstVals.length; r++) {
    if (String(gstVals[r][iInv] || '').trim() === String(excludeInvoice || '').trim()) continue;
    var value = String(gstVals[r][iDesc] || '').trim();
    if (!value) continue;
    var score = nameScore_(customer, String(gstVals[r][iName] || '').trim());
    if (score < 0.8) continue;
    var date = iDate >= 0 ? asDate_(gstVals[r][iDate]) : null;
    if (score > bestScore ||
        (score === bestScore && date && (!bestDate || date > bestDate)) ||
        (score === bestScore &&
         ((!date && !bestDate) || (date && bestDate && date.getTime() === bestDate.getTime())) &&
         r > bestRow)) {
      best = value;
      bestScore = score;
      bestDate = date;
      bestRow = r;
    }
  }
  return best;
}

/** '1 Year' -> 12, '6 Months' -> 6, '1 Month' -> 1. 0 if unparseable. */
function parseDuration_(text) {
  var months = 0, m;
  var re = /(\d+)\s*(year|month)/gi;
  while ((m = re.exec(String(text || ''))) !== null) {
    months += /^year/i.test(m[2]) ? parseInt(m[1], 10) * 12 : parseInt(m[1], 10);
  }
  return months;
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
