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
var COL_ADDRESS = 'Address';         // Optional legacy UPCOMING BILLS field, read by existing bill flows; not created during onboarding
var COL_USERID = 'User id';          // Optional legacy UPCOMING BILLS field, read by existing bill flows; not created during onboarding
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
  // New customer invoices default to Pending; blank status must never become Paid.
  row[gCols['Payment Status']] = String(v.payment_status || '').trim() || 'Pending';
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
  var desc = normalizeOnboardingDescription_(body.description);
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

  var map = loadJsonProperty_('DOCS_MAP');
  var manualMap = loadManualMap_();
  var requestedIdentity = customerIdentityContext_(sheetName);
  var conflictingMapKeys = [];
  [map, manualMap].forEach(function (currentMap) {
    Object.keys(currentMap).forEach(function (mappedName) {
      if (!requestedIdentity.matches(mappedName) || !currentMap[mappedName]) return;
      var mappedId = String(currentMap[mappedName]);
      var inactive = false;
      try {
        var mappedFile = DriveApp.getFileById(mappedId);
        inactive = mappedFile.isTrashed();
        if (isWithinCustomerBin_(mappedFile)) {
          conflictingMapKeys.push(mappedName + ' (still in Bin; permanently clean up before reusing this name)');
          return;
        }
      } catch (missingMappedFile) {
        inactive = true;
      }
      if (inactive) {
        delete currentMap[mappedName];
      } else {
        conflictingMapKeys.push(mappedName);
      }
    });
  });
  // Persist removal of trashed/missing/soft-deleted map entries.
  saveDocsMap_(map);
  saveManualMap_(manualMap);
  if (conflictingMapKeys.length) {
    throw new Error('A Google Doc is still actively mapped to this customer (' +
      conflictingMapKeys.join(', ') + '). Clean up the old deleted customer record before reusing the name.');
  }

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
      payment_status: 'Pending'
    });
    writeBillPeriod_(gst, ensureBillPeriodCol_(gst), firstGstRow - 1, periodStart, periodEnd);

    // Append only to existing core columns. Do not create or write Address/User id columns.
    var newRow = up.getLastRow() + 1;
    up.getRange(newRow, iSrNo + 1).setValue(nextUpcomingSrNo_(up, iSrNo));
    up.getRange(newRow, iName + 1).setValue(sheetName);
    var expCell = up.getRange(newRow, iExp + 1);
    expCell.setNumberFormat('dd/mm/yy');
    expCell.setValue(periodEnd);
    if (iDur >= 0) up.getRange(newRow, iDur + 1).setValue(durationText_(months));
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



