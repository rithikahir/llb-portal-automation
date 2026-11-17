/* =====================================================================
 * GST Bill Manager — PWA front-end (vanilla JS, no build step)
 *
 * Talks to the Apps Script Web App (gst-mobile-app/appscript/Code.js):
 *   ping / list_all / bill_customer / get_pdf / list_pdfs / make_pdf …
 * Each bill is created by editing that customer's own Google Doc in place,
 * exporting a PDF and archiving it on Drive (BILLS/<Customer>/). The phone
 * keeps no local copies — PDFs are streamed from Drive, viewable, shareable
 * (WhatsApp via the Android share sheet) and downloadable.
 * ===================================================================== */
'use strict';

/* ------------------------------------------------------------------ state */
var CFG_KEY = 'gst_cfg_v1';
var CACHE_KEY = 'gst_cache_v1';
var state = {
  cfg: null,          // {url, secret}
  data: null,         // list_all response {today, today_iso, customers:[]}
  savedAt: null,
  online: true,
  deferredInstall: null
};

/* ------------------------------------------------------------------ utils */
function $(sel, el) { return (el || document).querySelector(sel); }
function $all(sel, el) { return Array.prototype.slice.call((el || document).querySelectorAll(sel)); }
function esc(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function money(n) {
  var v = parseFloat(n);
  if (isNaN(v)) v = 0;
  return '₹' + v.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function todayISO() {
  var d = new Date();
  return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
}
function pad2(n) { return String(n).padStart(2, '0'); }

/** days from today until expire_iso (negative = already expired) */
function daysLeft(iso) {
  if (!iso) return null;
  var t = todayISO().split('-'), e = iso.split('-');
  var d1 = new Date(+t[0], +t[1] - 1, +t[2]);
  var d2 = new Date(+e[0], +e[1] - 1, +e[2]);
  return Math.round((d2 - d1) / 86400000);
}
function chipFor(iso) {
  var d = daysLeft(iso);
  if (d === null) return '<span class="chip grey">no date</span>';
  if (d < 0) return '<span class="chip red">EXPIRED ' + (-d) + 'd ago</span>';
  if (d === 0) return '<span class="chip red">EXPIRES TODAY</span>';
  if (d <= 7) return '<span class="chip amber">in ' + d + 'd</span>';
  if (d <= 30) return '<span class="chip green">in ' + d + 'd</span>';
  return '<span class="chip green">in ' + d + 'd</span>';
}
function statusChip(s) {
  s = String(s || '').toLowerCase();
  if (s.indexOf('paid') >= 0 && s.indexOf('un') < 0) return '<span class="chip green">Paid</span>';
  if (s.indexOf('pending') >= 0) return '<span class="chip amber">Pending</span>';
  return '<span class="chip grey">' + esc(s || '—') + '</span>';
}
/** "1 Year" -> 12, "6 Months" -> 6 (mirror of server parseDuration_) */
function parseDurationText(text) {
  var m = String(text || '').match(/(\d+)\s*(year|month)/i);
  if (!m) return 0;
  var n = parseInt(m[1], 10);
  return n && /^year/i.test(m[2]) ? n * 12 : n;
}
/** period end = (expire+1day advanced by months) - 1 day, clamped like server */
function periodEndISO(startISO, months) {
  var p = startISO.split('-');
  var start = new Date(+p[0], +p[1] - 1, +p[2]);
  var nextMonth = new Date(start.getFullYear(), start.getMonth() + months, 1);
  var dim = new Date(nextMonth.getFullYear(), nextMonth.getMonth() + 1, 0).getDate();
  nextMonth.setDate(Math.min(start.getDate(), dim));
  nextMonth.setDate(nextMonth.getDate() - 1);
  return nextMonth.getFullYear() + '-' + pad2(nextMonth.getMonth() + 1) + '-' + pad2(nextMonth.getDate());
}
function isoToDMY(iso) {
  if (!iso) return '—';
  var p = iso.split('-');
  return p[2] + '/' + p[1] + '/' + p[0];
}
function toast(msg, isErr) {
  var t = $('#toast');
  t.textContent = msg;
  t.className = 'toast show' + (isErr ? ' err' : '');
  clearTimeout(t._t);
  t._t = setTimeout(function () { t.className = 'toast' + (isErr ? ' err' : ''); }, 3200);
}
function busy(show, text) {
  var b = $('#busy');
  if (!b) return;
  $('.btxt', b).textContent = text || 'Working…';
  b.className = 'busy' + (show ? '' : ' hidden');
}

/* ------------------------------------------------------------- persistence */
function loadCfg() {
  try { state.cfg = JSON.parse(localStorage.getItem(CFG_KEY) || 'null'); } catch (e) { state.cfg = null; }
}
function saveCfg() { localStorage.setItem(CFG_KEY, JSON.stringify(state.cfg)); }
function saveCache() {
  try {
    localStorage.setItem(CACHE_KEY, JSON.stringify({ data: state.data, savedAt: state.savedAt }));
  } catch (e) { /* quota — ignore */ }
}
function loadCache() {
  try {
    var c = JSON.parse(localStorage.getItem(CACHE_KEY) || 'null');
    if (c && c.data) { state.data = c.data; state.savedAt = c.savedAt; }
  } catch (e) { }
}

/* ------------------------------------------------------------------ API */
function api(action, payload) {
  if (!state.cfg || !state.cfg.url) return Promise.reject(new Error('not configured'));
  var body = JSON.stringify(Object.assign({ action: action, secret: state.cfg.secret }, payload || {}));
  // Content-Type text/plain => "simple request", no CORS preflight needed
  return fetch(state.cfg.url, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: body
  }).then(function (r) {
    return r.json();
  }).then(function (j) {
    if (!j || j.ok === false) throw new Error((j && j.error) || 'server error');
    return j;
  });
}

/* --------------------------------------------------------- PDF utilities */
function b64ToBlob(b64) {
  var bin = atob(b64), arr = new Uint8Array(bin.length);
  for (var i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return new Blob([arr], { type: 'application/pdf' });
}
function downloadBlob(blob, name) {
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name || 'bill.pdf';
  document.body.appendChild(a);
  a.click();
  setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
}
function shareBlob(blob, name) {
  if (navigator.canShare && navigator.canShare({ files: [new File([blob], name, { type: 'application/pdf' })] })) {
    return navigator.share({ files: [new File([blob], name, { type: 'application/pdf' })], title: name })
      .catch(function () { });
  }
  downloadBlob(blob, name);
}

/* ------------------------------------------------------------ navigation */
var stack = [];   // [{fn, params}]
function go(fn, params) {
  stack.push({ fn: fn, params: params || {} });
  try { history.pushState({ d: stack.length }, ''); } catch (e) { }
  renderTop();
}
function back() {
  if (stack.length > 1) {
    stack.pop();
    renderTop();
    return true;
  }
  return false;
}
window.addEventListener('popstate', function () {
  if (back()) {
    // consumed; push a replacement state so the gesture works again
    try { history.pushState({ d: stack.length }, ''); } catch (e) { }
  }
});
function renderTop() {
  var top = stack[stack.length - 1];
  var app = $('#app');
  app.innerHTML = '';
  top.fn(app, top.params);
  window.scrollTo(0, 0);
}

/* ================================================================= SCREENS */

/* ------------------------------------------------------- setup / connect */
function renderSetup(app, retestOnly) {
  app.innerHTML =
    '<div class="hdr"><div class="hdr-row"><h1>📱 GST Bill Manager</h1></div>' +
    '<div class="sub">Connect to your GST sheet</div></div>' +
    '<div class="wrap">' +
    '<div class="card">' +
    '<div class="field"><label>Apps Script Web App URL (…/exec)</label>' +
    '<input id="c_url" type="url" placeholder="https://script.google.com/macros/s/…/exec" value="' + esc(state.cfg ? state.cfg.url : '') + '"></div>' +
    '<div class="field"><label>Shared secret</label>' +
    '<input id="c_secret" type="password" placeholder="appscript_secret" value="' + esc(state.cfg ? state.cfg.secret : '') + '"></div>' +
    '<button class="btn primary" id="c_save">Connect</button>' +
    '</div>' +
    '<div class="card"><b>Where do I get these?</b><div class="meta-line" style="margin-top:6px">' +
    'Deploy <b>gst-mobile-app\\appscript\\Code.js</b> as a Web App ' +
    '(Extensions → Apps Script in your GST sheet → Deploy → Web app, ' +
    'execute as <b>Me</b>, access <b>Anyone</b>). The URL ends in <b>/exec</b>. ' +
    'The secret is the <b>appscript_secret</b> from your .env file.' +
    '</div></div></div>';
  $('#c_save').onclick = function () {
    var url = $('#c_url').value.trim().replace(/\?.*$/, '');
    var secret = $('#c_secret').value.trim();
    if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) { toast('URL must end in /exec', true); return; }
    if (!secret) { toast('Enter the shared secret', true); return; }
    state.cfg = { url: url, secret: secret };
    busy(true, 'Testing connection…');
    api('ping').then(function () {
      saveCfg();
      toast('Connected ✓');
      stack = [{ fn: renderTab, params: { tab: 'today' } }];
      renderTop();
      refreshData(true);
    }).catch(function (e) {
      busy(false);
      toast('Connection failed: ' + e.message, true);
    });
  };
}

/* --------------------------------------------------------------- tabs */
var TABS = { today: 'Home', customers: 'Customers', expiries: 'Expiries', settings: 'Settings' };

function renderTab(app, params) {
  var tab = params.tab;
  app.innerHTML =
    '<div class="hdr"><div class="hdr-row">' +
    '<h1>' + (tab === 'today' ? 'GST Bills' : TABS[tab]) + '</h1>' +
    '<span style="display:flex;gap:8px;align-items:center">' +
    '<button class="iconbtn" id="addcust" title="Add customer" style="width:34px;height:34px;font-size:20px;line-height:1">＋</button>' +
    '<button class="iconbtn' + (state._refreshing ? ' spin' : '') + '" id="refresh" title="Refresh">' + svgRefresh() + '</button>' +
    '</span>' +
    '</div><div class="sub" id="hdrsub">' + hdrSub(tab) + '</div></div>' +
    (state.online ? '' : '<div class="banner">⚠ Offline — showing last saved data</div>') +
    '<div id="tabbody"></div>' +
    '<nav class="nav">' +
    navBtn('today', '🏠', 'Home', tab, '') +
    navBtn('customers', '👥', 'Customers', tab) +
    navBtn('expiries', '⏳', 'Expiries', tab) +
    navBtn('settings', '⚙️', 'Settings', tab) +
    '</nav>';
  $('#refresh').onclick = function () { refreshData(true); };
  $('#addcust').onclick = function () { go(renderAddCustomer, {}); };
  $all('.nav button').forEach(function (b) {
    b.onclick = function () {
      if (b.dataset.tab === tab) return;
      stack = [{ fn: renderTab, params: { tab: b.dataset.tab } }];
      renderTop();
    };
  });
  var body = $('#tabbody');
  if (tab === 'today') renderToday(body);
  else if (tab === 'customers') renderCustomersTab(body);
  else if (tab === 'expiries') renderExpiriesTab(body);
  else renderSettings(body);
}
function navBtn(tab, icon, label, active, extra) {
  return '<button data-tab="' + tab + '" class="' + (tab === active ? 'active' : '') + '">' +
    '<span class="ni">' + icon + '</span>' + label + (extra || '') + '</button>';
}
function hdrSub(tab) {
  if (!state.data) return 'Not loaded yet';
  var cs = state.data.customers;
  var today = cs.filter(function (c) { return daysLeft(c.expire_iso) === 0; }).length;
  var overdue = cs.filter(function (c) { var d = daysLeft(c.expire_iso); return d !== null && d < 0; }).length;
  var when = state.savedAt ? new Date(state.savedAt).toLocaleString() : '';
  if (tab === 'today') return today + ' expiring today · ' + overdue + ' overdue · ' + cs.length + ' customers' + (when ? ' · ' + when : '');
  return cs.length + ' customers' + (when ? ' · updated ' + when : '');
}
function svgRefresh() {
  return '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M21 12a9 9 0 1 1-2.6-6.4"/><path d="M21 3v6h-6"/></svg>';
}

/* ------------------------------------------------------------ data refresh */
function refreshData(force) {
  if (!state.cfg) { stack = [{ fn: renderSetup }]; renderTop(); return; }
  if (state._refreshing) return;
  state._refreshing = true;
  var btn = $('#refresh');
  if (btn) btn.className = 'iconbtn spin';
  api('list_all').then(function (d) {
    state.data = d;
    state.savedAt = Date.now();
    state.online = true;
    saveCache();
    toast('Updated ✓');
  }).catch(function (e) {
    state.online = false;
    if (force) toast('Refresh failed: ' + e.message, true);
  }).finally(function () {
    state._refreshing = false;
    renderTop();
  });
}

/* ---------------------------------------------------------------- HOME */
function customerCard(c, opts) {
  opts = opts || {};
  var last = c.last_invoice;
  return '<div class="card tap" data-name="' + esc(c.name) + '">' +
    '<div class="row-flex">' +
    '<div class="grow">' +
    '<p class="cust-name">' + esc(c.name) + '</p>' +
    '<div class="meta-line">' +
    (last ? 'Last: ' + esc(last.invoice_no) + ' · ' + esc(last.invoice_date) + (last.status ? ' · ' : '') : 'No invoice yet') +
    (last && last.status ? statusChip(last.status) : '') +
    '</div></div>' +
    '<div class="right">' + chipFor(c.expire_iso) +
    (last ? '<div class="amount" style="margin-top:5px">' + money(last.total) + '</div>' : '') +
    '</div></div></div>';
}
function bindCustomerCards(root) {
  $all('.card.tap[data-name]', root).forEach(function (el) {
    el.onclick = function () { go(renderCustomerDetail, { name: el.dataset.name }); };
  });
}

function renderToday(body) {
  if (!state.data) {
    body.innerHTML = '<div class="empty"><div class="big">📭</div>No data yet.<br><br>' +
      '<button class="btn primary" onclick="refreshData(true)">Load now</button></div>';
    return;
  }
  var cs = state.data.customers;
  var today = cs.filter(function (c) { return daysLeft(c.expire_iso) === 0; });
  var overdue = cs.filter(function (c) { var d = daysLeft(c.expire_iso); return d !== null && d < 0; });
  var week = cs.filter(function (c) { var d = daysLeft(c.expire_iso); return d !== null && d > 0 && d <= 7; });
  var html = '';

  if (today.length) {
    html += '<div class="section-title" style="color:var(--danger)">🔴 Expires today (' + today.length + ')</div>' +
      '<div class="card" style="background:var(--danger-soft);border-color:#fecaca">' +
      '<div class="meta-line" style="color:var(--danger)"><b>Bill these customers now</b> — their plan ends today.</div></div>' +
      today.map(function (c) { return customerCard(c); }).join('');
  }
  if (overdue.length) {
    html += '<div class="section-title" style="color:var(--danger)">⚠ Overdue (' + overdue.length + ')</div>' +
      overdue.map(function (c) { return customerCard(c); }).join('');
  }
  if (week.length) {
    html += '<div class="section-title">Next 7 days (' + week.length + ')</div>' +
      week.map(function (c) { return customerCard(c); }).join('');
  }
  if (!html) html = '<div class="empty"><div class="big">✅</div>Nothing due today.<br>All ' + cs.length + ' customers are paid up ahead.</div>';
  body.innerHTML = '<div class="wrap">' + html + '</div>';
  bindCustomerCards(body);
}

/* ------------------------------------------------------------ CUSTOMERS */
function renderCustomersTab(body) {
  body.innerHTML = '<div class="wrap">' +
    '<div class="search"><span class="si">🔍</span>' +
    '<input id="q" placeholder="Search name, user id, GSTIN…" value="' + esc(state._q || '') + '"></div>' +
    '<div id="clist"></div></div>';
  var q = $('#q');
  q.oninput = function () { state._q = q.value; fillCustomerList($('#clist'), q.value); };
  fillCustomerList($('#clist'), state._q || '');
}
function fillCustomerList(root, query) {
  if (!state.data) {
    root.innerHTML = '<div class="empty"><div class="big">📭</div>No data yet.</div>';
    return;
  }
  var q = String(query || '').trim().toLowerCase();
  var cs = state.data.customers.filter(function (c) {
    if (!q) return true;
    return (c.name + ' ' + c.userid + ' ' + c.gstin).toLowerCase().indexOf(q) >= 0;
  });
  root.innerHTML = cs.length
    ? cs.map(function (c) { return customerCard(c); }).join('')
    : '<div class="empty"><div class="big">🔎</div>No customers match.</div>';
  bindCustomerCards(root);
}

/* ------------------------------------------------------------- EXPIRIES */
function renderExpiriesTab(body) {
  if (!state.data) {
    body.innerHTML = '<div class="wrap"><div class="empty"><div class="big">📭</div>No data yet.</div></div>';
    return;
  }
  var cs = state.data.customers.filter(function (c) { return c.expire_iso; });
  var groups = [
    ['🔴 Expired', function (d) { return d < 0; }, 'var(--danger)'],
    ['🔴 Expires today', function (d) { return d === 0; }, 'var(--danger)'],
    ['🟠 Next 7 days', function (d) { return d > 0 && d <= 7; }, 'var(--warn)'],
    ['🟡 Next 30 days', function (d) { return d > 7 && d <= 30; }, 'var(--warn)'],
    ['🟢 Later', function (d) { return d > 30; }, 'var(--ok)']
  ];
  var html = '<div class="wrap">';
  groups.forEach(function (g) {
    var list = cs.filter(function (c) { var d = daysLeft(c.expire_iso); return d !== null && g[1](d); });
    if (!list.length) return;
    html += '<div class="section-title" style="color:' + g[2] + '">' + g[0] + ' (' + list.length + ')</div>';
    html += list.map(function (c) {
      return '<div class="card tap" data-name="' + esc(c.name) + '"><div class="row-flex">' +
        '<div class="grow"><p class="cust-name">' + esc(c.name) + '</p>' +
        '<div class="meta-line">' + esc(c.expire) + (c.duration ? ' · ' + esc(c.duration) : '') + '</div></div>' +
        '<div class="right">' + chipFor(c.expire_iso) + '</div></div></div>';
    }).join('');
  });
  if (!html.trim()) html += '<div class="empty"><div class="big">✅</div>No expiry data.</div>';
  body.innerHTML = html + '</div>';
  bindCustomerCards(body);
}

/* ------------------------------------------------------ CUSTOMER DETAIL */
function renderCustomerDetail(app, params) {
  var c = null;
  if (state.data) {
    c = state.data.customers.filter(function (x) { return x.name === params.name; })[0];
  }
  app.innerHTML =
    '<div class="overlay"><div class="ohdr">' +
    '<button class="back" id="bk">←</button><h2>' + esc(params.name) + '</h2></div>' +
    (c ? detailBody(c) : '<div class="wrap"><div class="empty"><div class="big">❓</div>Customer not found. Refresh data.</div></div>') +
    '</div>';
  $('#bk').onclick = back;
  if (!c) return;
  var bb = $('#billnow');
  if (bb) bb.onclick = function () { go(renderBillForm, { name: c.name }); };
  // every archived PDF lives on Drive — tapping opens it from there
  $all('.pdfbtn[data-inv]:not(.paybtn)').forEach(function (b) {
    b.onclick = function () { openPdfViewer(b.dataset.inv); };
  });
  // payment status editor — one per invoice row
  $all('.paybtn[data-inv]').forEach(function (b) {
    b.onclick = function () {
      var inv = c.invoices.filter(function (x) { return x.invoice_no === b.dataset.inv; })[0];
      if (inv) go(renderPaymentEdit, { c: c, inv: inv });
    };
  });
  // bill editor — one per invoice row
  $all('.ebtn[data-inv]').forEach(function (b) {
    b.onclick = function () {
      var inv = c.invoices.filter(function (x) { return x.invoice_no === b.dataset.inv; })[0];
      if (inv) go(renderBillEdit, { c: c, inv: inv });
    };
  });
}
function detailBody(c) {
  var last = c.last_invoice;
  var info =
    '<div class="card"><div class="row-flex">' +
    '<div class="grow"><div class="meta-line">Plan expiry</div>' +
    '<div style="font-weight:800;font-size:17px">' + esc(c.expire || '—') + '</div>' +
    '<div style="margin-top:5px">' + chipFor(c.expire_iso) + '</div></div>' +
    '<div class="right"><div class="meta-line">Duration</div><div style="font-weight:700">' + esc(c.duration || '—') + '</div></div>' +
    '</div>' +
    '<div class="meta-line" style="margin-top:10px">' +
    (c.userid ? 'User id: <b>' + esc(c.userid) + '</b> · ' : '') +
    (c.gstin ? 'GSTIN: <b>' + esc(c.gstin) + '</b>' : '') +
    '</div>' +
    (c.bill_file ? '<div class="small" style="margin-top:4px">Bill file: ' + esc(c.bill_file) + '</div>' : '') +
    (c.address ? '<div class="small" style="margin-top:4px;white-space:pre-line">📍 ' + esc(c.address) + '</div>' : '') +
    '<div class="btn-row"><button class="btn primary" id="billnow" style="margin-top:6px">🧾 Update Bill / Bill Now</button></div>' +
    '</div>';

  var hist = '<div class="section-title">Invoice history (' + c.invoices.length + ')</div>';
  if (!c.invoices.length) {
    hist += '<div class="card"><div class="meta-line">No GST invoices found for this customer yet.</div></div>';
  } else {
    hist += '<div class="card">' + c.invoices.map(function (inv) {
      return '<div class="inv"><div class="grow">' +
        '<div class="no">' + esc(inv.invoice_no) + '</div>' +
        '<div class="date">' + esc(inv.invoice_date) + ' · ' + esc(inv.desc || '') + '</div>' +
        '<div class="date">' + statusChip(inv.status) +
        (inv.payment_date ? ' · paid ' + esc(inv.payment_date) : '') + '</div></div>' +
        '<div class="right"><div class="amount">' + money(inv.total) + '</div>' +
        '<button class="pdfbtn" data-inv="' + esc(inv.invoice_no) + '">PDF 📄</button>' +
        '<button class="pdfbtn paybtn" data-inv="' + esc(inv.invoice_no) + '">💳 ' + esc(inv.status || '?') + '</button>' +
        '<button class="pdfbtn ebtn" data-inv="' + esc(inv.invoice_no) + '">✏️ Edit</button></div></div>';
    }).join('') + '</div>';
  }
  return '<div class="wrap">' + info + hist + '</div>';
}

/* --------------------------------------------------------------- BILLING */
/* ------------------------------------------------- payment status editor */
function renderPaymentEdit(app, params) {
  var c = params.c, inv = params.inv;
  var cur = (inv.status || 'Pending').toLowerCase() === 'paid' ? 'Paid' : 'Pending';
  // prefill: existing payment date if parseable, else today (yyyy-mm-dd)
  var preISO = '';
  if (inv.payment_date) { var d = new Date(inv.payment_date);
    if (!isNaN(d)) preISO = d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }
  app.innerHTML =
    '<div class="overlay"><div class="ohdr">' +
    '<button class="back" id="bk">←</button><h2>💳 Payment · ' + esc(c.name) + '</h2></div>' +
    '<div class="wrap">' +
    '<div class="summary"><div class="line"><span>Invoice</span><b>' + esc(inv.invoice_no) + '</b></div>' +
    '<div class="line"><span>Total</span><b>' + money(inv.total) + '</b></div>' +
    '<div class="line"><span>Current status</span><b>' + esc(inv.status || '?') +
    (inv.payment_date ? ' · ' + esc(inv.payment_date) : '') + '</b></div></div>' +
    '<div class="card">' +
    '<div class="field"><label>Payment status</label><select id="p_status">' +
    '<option value="Paid"' + (cur === 'Paid' ? ' selected' : '') + '>Paid</option>' +
    '<option value="Pending"' + (cur === 'Pending' ? ' selected' : '') + '>Pending</option>' +
    '</select></div>' +
    '<div class="field" id="p_datebox"><label>Payment date</label>' +
    '<input id="p_date" type="date" value="' + (preISO || '') + '"></div>' +
    '</div>' +
    '<button class="btn primary" id="doPay">💾 Save to sheet</button>' +
    '<div style="height:20px"></div></div></div>';

  $('#bk').onclick = back;
  var syncDate = function () {
    $('#p_datebox').style.display = $('#p_status').value === 'Paid' ? 'block' : 'none';
  };
  $('#p_status').onchange = syncDate;
  syncDate();
  if (!preISO) $('#p_date').value = new Date().getFullYear() + '-' + pad2(new Date().getMonth() + 1) + '-' + pad2(new Date().getDate());

  $('#doPay').onclick = function () {
    var st = $('#p_status').value;
    var pd = $('#p_date').value;
    if (st === 'Paid' && !pd) { toast('Pick the payment date', true); return; }
    busy(true, 'Updating sheet…');
    api('update_payment', {
      customer: c.name, invoice_no: inv.invoice_no,
      payment_status: st, payment_date: pd
    }).then(function (res) {
      busy(false);
      toast('Saved ✓ ' + res.invoice_no + ' → ' + res.payment_status +
            (res.payment_date ? ' (' + res.payment_date + ')' : ''));
      back();
      refreshData(false);
    }).catch(function (e) {
      busy(false);
      toast('Update failed: ' + e.message, true);
    });
  };
}

/* --------------------------------------------------------- bill editor --- */
function renderBillEdit(app, params) {
  var c = params.c, inv = params.inv;
  var toISO = function (d) {
    var dt = new Date(d);
    return isNaN(dt) ? '' : dt.getFullYear() + '-' + pad2(dt.getMonth() + 1) + '-' + pad2(dt.getDate());
  };
  app.innerHTML =
    '<div class="overlay"><div class="ohdr">' +
    '<button class="back" id="bk">←</button><h2>✏️ Edit · ' + esc(inv.invoice_no) + '</h2></div>' +
    '<div class="wrap">' +
    '<div id="ebbody"><div class="empty"><div class="big">📄</div>Loading bill…</div></div>' +
    '<div style="height:20px"></div></div></div>';
  $('#bk').onclick = back;

  api('bill_edit_prefill', { customer: c.name, invoice_no: inv.invoice_no })
    .then(function (pf) {
      if (!pf || pf.ok === false) throw new Error('no prefill');
      var latestNote = pf.is_latest
        ? '<div class="card"><div class="meta-line">📌 This is the <b>latest</b> bill — the doc shows it now; editing updates the doc, PDF and expiry.</div></div>'
        : '<div class="card"><div class="meta-line">📌 This is an <b>older</b> bill — the doc will temporarily switch to it for the PDF, then restore the latest bill automatically. Expiry on the sheet stays with the latest bill. Bill period is required (check the current PDF for it).</div></div>';
      $('#ebbody').innerHTML =
        latestNote +
        '<div class="card">' +
        '<div class="field"><label>Invoice no (fixed)</label><input id="e_inv" value="' + esc(inv.invoice_no) + '" disabled></div>' +
        '<div class="field"><label>Buyer name (printed on doc)</label><input id="e_name" value="' + esc(pf.name || c.name) + '"></div>' +
        '<div class="field"><label>Address</label><input id="e_addr" value="' + esc(pf.address || '') + '"></div>' +
        '<div class="grid2">' +
        '<div class="field"><label>GSTIN / UIN (optional)</label><input id="e_gstin" value="' + esc(pf.gstin || '') + '"></div>' +
        '<div class="field"><label>User id</label><input id="e_userid" value="' + esc(pf.userid || '') + '"></div>' +
        '</div>' +
        '<div class="field"><label>Invoice date</label><input id="e_date" type="date" value="' + esc(pf.invoice_date || '') + '"></div>' +
        '<div class="grid2">' +
        '<div class="field"><label>Bill period start</label><input id="e_pstart" type="date" value="' + esc(pf.period_start || '') + '"></div>' +
        '<div class="field"><label>Bill period end</label><input id="e_pend" type="date" value="' + esc(pf.period_end || '') + '"></div>' +
        '</div>' +
        '<div class="field"><label>Service description (exact PDF text)</label>' +
        '<textarea id="e_desc">' + esc(pf.desc || '') + '</textarea></div>' +
        '<div class="field"><label>Amount ₹ (CGST/SGST/total recompute)</label>' +
        '<input id="e_amount" type="number" inputmode="decimal" step="0.01" min="0" value="' + esc(pf.amount || '') + '"></div>' +
        '</div>' +
        '<button class="btn primary" id="e_save">💾 Save edits &amp; regenerate PDF</button>';

      $('#e_save').onclick = function () {
        var body = {
          customer: c.name,
          invoice_no: inv.invoice_no,
          name: $('#e_name').value.trim(),
          address: $('#e_addr').value.trim(),
          gstin: $('#e_gstin').value.trim(),
          userid: $('#e_userid').value.trim(),
          invoice_date: $('#e_date').value,
          period_start: $('#e_pstart').value,
          period_end: $('#e_pend').value,
          desc: $('#e_desc').value,
          amount: parseFloat($('#e_amount').value)
        };
        if (!(body.amount > 0)) { toast('Enter a valid amount', true); return; }
        if (!body.invoice_date) { toast('Pick the invoice date', true); return; }
        if (!body.period_start || !body.period_end) {
          toast('Bill period (start + end) is required — check the current PDF', true); return;
        }
        var cg = Math.round(body.amount * 9) / 100;
        var tot = Math.round((body.amount + cg + cg) * 100) / 100;
        if (!confirm('Edit bill ' + inv.invoice_no + ' of ' + c.name + '?\n\n' +
            'Invoice date: ' + body.invoice_date + '\n' +
            'Period: ' + body.period_start + ' → ' + body.period_end + '\n' +
            'Amount: ' + money(body.amount) + ' → total ' + money(tot) + '\n\n' +
            'The Google Doc and the archived PDF for THIS bill are updated.' +
            (pf.is_latest ? ' Expiry moves to the new period end.' :
              ' The doc keeps the latest bill afterwards.'))) return;
        busy(true, 'Editing bill… regenerating PDF…');
        api('edit_bill', body).then(function (res) {
          busy(false);
          toast('Bill updated ✓ ' + res.invoice_no);
          back(); back();               // editor + detail screen both refresh
          refreshData(false);
        }).catch(function (e) {
          busy(false);
          toast('Edit failed: ' + e.message, true);
        });
      };
    })
    .catch(function (e) {
      $('#ebbody').innerHTML = '<div class="empty"><div class="big">⚠️</div>' +
        esc(e.message) + '</div>';
    });
}

/* ------------------------------------------------------ onboarding (new) */
function renderAddCustomer(app) {
  app.innerHTML =
    '<div class="overlay"><div class="ohdr">' +
    '<button class="back" id="bk">←</button><h2>＋ New customer</h2></div>' +
    '<div class="wrap">' +

    '<div class="summary" id="sumbox"></div>' +

    '<div class="card">' +
    '<div class="field"><label>Customer name (printed on doc) *</label>' +
    '<input id="a_name" placeholder="e.g. Sharma Traders"></div>' +
    '<div class="field"><label>Sheet record name (GST RECORDS "Customer Name")</label>' +
    '<input id="a_sheetname" placeholder="same as doc name unless different" data-auto="1"></div>' +
    '<div class="field"><label>Doc file name (in Drive)</label>' +
    '<input id="a_docname" placeholder="same as doc name unless different" data-auto="1"></div>' +
    '<div class="field"><label>Address * (printed below the name)</label>' +
    '<textarea id="a_address" placeholder="Street, area, Pune - PIN"></textarea></div>' +
    '<div class="grid2">' +
    '<div class="field"><label>Plan amount (₹) *</label>' +
    '<input id="a_amount" type="number" inputmode="decimal" step="0.01" min="0" placeholder="e.g. 500"></div>' +
    '<div class="field"><label>Duration (months) *</label><select id="a_months">' +
    [1, 3, 6, 12].map(function (m) {
      return '<option value="' + m + '"' + (m === 1 ? ' selected' : '') + '>' + m + '</option>';
    }).join('') + '</select></div>' +
    '</div>' +
    '<div class="field"><label>Bill period starts *</label>' +
    '<input id="a_start" type="date"></div>' +
    '<div class="field"><label>Plan description *</label>' +
    '<textarea id="a_desc">HIGH SPEED 100 MBPS\nTRULY UNLIMITED 12 MONTHS PLAN.</textarea></div>' +
    '<div class="grid2">' +
    '<div class="field"><label>GSTIN / UIN (optional — hidden if empty)</label><input id="a_gstin" placeholder="27ABCDE1234F1Z5"></div>' +
    '<div class="field"><label>User id *</label><input id="a_userid" placeholder="e.g. SHARMA01"></div>' +
    '</div>' +
    '<div class="field"><label>Payment mode</label><input id="a_mode" value="Online"></div>' +
    '</div>' +

    '<div class="card"><div class="meta-line">📌 On confirm: invoice no is auto-assigned, the customer\'s <b>own Google Doc is created</b> (same layout as your other bills), the first bill PDF is archived on Drive under <b>BILLS/&lt;name&gt;</b>, a <b>GST RECORDS</b> row (Paid) is added, and the customer appears in <b>UPCOMING BILLS</b> with the new expiry.</div></div>' +

    '<button class="btn primary" id="doAdd">🧾 Create customer &amp; first bill</button>' +
    '<div style="height:20px"></div>' +
    '</div></div>';

  $('#bk').onclick = back;
  // the two derived name fields mirror the doc name until edited by hand
  ['#a_sheetname', '#a_docname'].forEach(function (sel) {
    $(sel).oninput = function () { this.dataset.auto = '0'; };
  });
  $('#a_name').oninput = function () {
    ['#a_sheetname', '#a_docname'].forEach(function (sel) {
      var el = $(sel);
      if (el.dataset.auto === '1') el.value = this.value;
    }.bind(this));
  };
  var todayISO = function () {
    var d = new Date();
    return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
  };
  $('#a_start').value = todayISO();
  var recalc = function () {
    var m = parseInt($('#a_months').value, 10) || 1;
    var amt = parseFloat($('#a_amount').value) || 0;
    var sParts = ($('#a_start').value || todayISO()).split('-');
    var p = new Date(+sParts[0], +sParts[1] - 1, +sParts[2]);
    var e = new Date(p); e.setMonth(e.getMonth() + m); e.setDate(e.getDate() - 1);
    function dmy(d) { return pad2(d.getDate()) + '/' + pad2(d.getMonth() + 1) + '/' + String(d.getFullYear()).slice(2); }
    var cgst = Math.round(amt * 9) / 100, sgst = cgst, tot = Math.round((amt + cgst + sgst) * 100) / 100;
    $('#sumbox').innerHTML =
      '<div class="line"><span>Bill period</span><b>' + (amt > 0 ? dmy(p) + ' → ' + dmy(e) : '—') + '</b></div>' +
      '<div class="line"><span>Expiry</span><b>' + (amt > 0 ? dmy(e) : '—') + '</b></div>' +
      '<div class="line"><span>Basic / CGST / SGST</span><b>' + money(amt) + ' / ' + money(cgst) + ' / ' + money(sgst) + '</b></div>' +
      '<div class="line"><span>Invoice total</span><b class="big">' + money(tot) + '</b></div>' +
      '<div class="line"><span>Invoice no</span><b>auto (next in FY)</b></div>';
  };
  ['#a_months', '#a_amount'].forEach(function (s) { $(s).oninput = recalc; $(s).onchange = recalc; });
  recalc();

  $('#doAdd').onclick = function () {
    var nm = $('#a_name').value.trim();
    var amt = parseFloat($('#a_amount').value);
    var m = parseInt($('#a_months').value, 10) || 1;
    var ad = $('#a_address').value.trim();
    var uid = $('#a_userid').value.trim();
    var ds = $('#a_desc').value.trim();
    if (!nm) { toast('Enter the customer name', true); return; }
    if (!ad) { toast('Enter the address (printed below the name)', true); return; }
    if (!uid) { toast('Enter the user id', true); return; }
    if (!ds) { toast('Enter the plan description', true); return; }
    if (!(amt > 0)) { toast('Enter a valid amount', true); return; }
    if (!confirm('Onboard ' + nm + '?\n\nFirst invoice: next number in FY\nAmount: ' + money(amt) + ' (+9% CGST + 9% SGST), ' + m + ' month(s)\nPeriod: ' + $('#a_start').value + ' → ' + ($('#a_months').value) + ' month(s)\n\nCreates their Google Doc, archives the PDF, and adds them to the sheet.')) return;
    busy(true, 'Creating customer… making their doc…');
    api('add_customer', {
      name: nm,
      sheet_name: $('#a_sheetname').value.trim(),
      doc_name: $('#a_docname').value.trim(),
      amount: amt,
      duration_months: m,
      period_start: $('#a_start').value,
      description: ds,
      gstin: $('#a_gstin').value.trim(),
      userid: uid,
      address: ad,
      payment_mode: $('#a_mode').value.trim()
    }).then(function (res) {
      busy(false);
      toast('Customer created ✓ ' + res.invoice_no);
      go(renderBillSuccess, { result: res });
      refreshData(false);
    }).catch(function (e) {
      busy(false);
      toast('Onboarding failed: ' + e.message, true);
    });
  };
}

function renderBillForm(app, params) {
  var c = state.data.customers.filter(function (x) { return x.name === params.name; })[0];
  if (!c) { toast('Customer not found', true); back(); return; }
  var last = c.last_invoice;
  var defMonths = parseDurationText(c.duration) || 12;
  var defAmount = last && last.amount > 0 ? last.amount : '';
  var defDesc = last ? last.desc : '';
  var defMode = last && last.mode ? last.mode : 'Cash';

  app.innerHTML =
    '<div class="overlay"><div class="ohdr">' +
    '<button class="back" id="bk">←</button><h2>Bill · ' + esc(c.name) + '</h2></div>' +
    '<div class="wrap">' +

    '<div class="summary" id="sumbox"></div>' +

    '<div class="card">' +
    '<div class="grid2">' +
    '<div class="field"><label>Duration (months)</label><select id="f_months">' +
    [1, 3, 6, 12, 13, 24].map(function (m) {
      return '<option value="' + m + '"' + (m === defMonths ? ' selected' : '') + '>' + m + '</option>';
    }).join('') + '</select></div>' +
    '<div class="field"><label>Amount (₹)</label><input id="f_amount" type="number" inputmode="decimal" step="0.01" min="0" value="' + esc(defAmount) + '"></div>' +
    '</div>' +
    '<div class="field"><label>Service description</label><textarea id="f_desc">' + esc(defDesc) + '</textarea></div>' +
    '<div class="grid2">' +
    '<div class="field"><label>Payment mode</label><input id="f_mode" value="' + esc(defMode) + '"></div>' +
    '<div class="field"><label>User id</label><input id="f_userid" value="' + esc(c.userid || '') + '" placeholder="e.g. HINDUSTAN"></div>' +
    '</div>' +
    '<div class="field"><label>Buyer address (saved to sheet for next time)</label>' +
    '<textarea id="f_address" placeholder="Street, area, Pune - PIN">' + esc(c.address || '') + '</textarea></div>' +
    '</div>' +

    '<div class="card">' +
    '<label style="display:flex;align-items:center;gap:8px;padding:4px 0"><input type="radio" name="ovmode" id="ov_off" checked style="width:18px;height:18px">' +
    '<span>Invoice date = today</span></label>' +
    '<label style="display:flex;align-items:center;gap:8px;padding:4px 0"><input type="radio" name="ovmode" id="ov_on" style="width:18px;height:18px">' +
    '<span>Bill early — override invoice date</span></label>' +
    '<div class="field" id="ov_box" style="display:none;margin-top:6px">' +
    '<label>Invoice date (must be after expiry)</label>' +
    '<input type="date" id="f_invdate" min="' + minISO + '">' +
    '</div>' +
    '</div>' +

    '</div>' +

    '<div class="card"><div class="meta-line">📌 On confirm: next global invoice no is assigned, a row is added to <b>GST RECORDS</b> (Payment Pending), expiry moves to the new period end, the tab is re-sorted, and the bill <b>PDF is generated and saved on this phone</b> (no Drive copy).</div></div>' +

    '<button class="btn primary" id="doBill">🧾 Update Bill &amp; Advance Expiry</button>' +
    '<div style="height:20px"></div>' +
    '</div></div>';

  $('#bk').onclick = back;
  // expiry / invoice-date guard state
  var expParts = (c.expire_iso || '').split('-');
  var expD = expParts.length === 3 ? new Date(+expParts[0], +expParts[1] - 1, +expParts[2]) : null;
  var minISO = '';
  if (expD) { var m1 = new Date(expD); m1.setDate(m1.getDate() + 1);
    minISO = m1.getFullYear() + '-' + pad2(m1.getMonth() + 1) + '-' + pad2(m1.getDate()); }
  var today0 = new Date(); today0.setHours(0, 0, 0, 0);
  var billable = expD && expD.getTime() < today0.getTime();   // expiry strictly before today
  $('#ov_on').onchange = $('#ov_off').onchange = function () {
    $('#ov_box').style.display = $('#ov_on').checked ? 'block' : 'none';
    recalc();
  };
  var recalc = function () {
    var m = parseInt($('#f_months').value, 10) || 12;
    var amt = parseFloat($('#f_amount').value) || 0;
    var startISO = c.expire_iso;
    var endISO = startISO ? periodEndISO(startISO, m) : null;
    var cgst = Math.round(amt * 9) / 100, sgst = cgst, tot = Math.round((amt + cgst + sgst) * 100) / 100;
    $('#sumbox').innerHTML =
      '<div class="line"><span>Current expiry</span><b>' + isoToDMY(c.expire_iso) + '</b></div>' +
      '<div class="line"><span>New bill period</span><b>' + (startISO ? isoToDMY(periodEndPlus1(startISO)) + ' → ' + isoToDMY(endISO) : '—') + '</b></div>' +
      '<div class="line"><span>New expiry</span><b>' + (endISO ? isoToDMY(endISO) : '—') + '</b></div>' +
      '<div class="line"><span>Basic / CGST / SGST</span><b>' + money(amt) + ' / ' + money(cgst) + ' / ' + money(sgst) + '</b></div>' +
      '<div class="line"><span>Invoice total</span><b class="big">' + money(tot) + '</b></div>' +
      '<div class="line"><span>Invoice no</span><b>auto (next in FY)</b></div>' +
      '<div class="line"><span>Invoice date</span><b>' + ($('#ov_on') && $('#ov_on').checked && $('#f_invdate').value ? isoToDMY($('#f_invdate').value) + ' (override)' : 'today') + '</b></div>';
  };
  function periodEndPlus1(iso) { var p = iso.split('-'); var d = new Date(+p[0], +p[1] - 1, +p[2]); d.setDate(d.getDate() + 1); return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }
  ['#f_months', '#f_amount', '#f_invdate'].forEach(function (s) { $(s).oninput = recalc; $(s).onchange = recalc; });
  // Service description = EXACT text the doc (and therefore the PDF) shows
  api('bill_prefill', { customer: c.name }).then(function (res) {
    var el = $('#f_desc');
    if (res && res.desc && el && !el.dataset.touched) el.value = res.desc;
  }).catch(function () { /* keep the sheet desc */ });
  $('#f_desc').oninput = function () { this.dataset.touched = '1'; };
  recalc();

  $('#doBill').onclick = function () {
    var m = parseInt($('#f_months').value, 10) || 12;
    var amt = parseFloat($('#f_amount').value);
    if (!(amt > 0)) { toast('Enter a valid amount', true); return; }
    // invoice-date guard: normal billing only after expiry; early billing
    // requires the override radio + a date strictly after the expiry
    var invDate = '';
    if ($('#ov_on').checked) {
      invDate = $('#f_invdate').value;
      if (!invDate) { toast('Pick the override invoice date', true); return; }
      var picked = new Date(invDate + 'T00:00:00');
      if (expD && picked <= expD) {
        toast('Invoice date must be AFTER the expiry ' + isoToDMY(c.expire_iso), true); return;
      }
    } else if (!billable) {
      toast(c.name + ' is covered until ' + isoToDMY(c.expire_iso) +
        ' — billable from ' + isoToDMY(minISO) + '. To bill early, use the override.', true);
      return;
    }
    var endISO = c.expire_iso ? periodEndISO(c.expire_iso, m) : null;
    if (!confirm('Bill ' + c.name + '?\n\nInvoice: next number in FY\nInvoice date: ' + (invDate ? isoToDMY(invDate) : 'today') + '\nAmount: ' + money(amt) + ' (+9% CGST + 9% SGST)\nNew expiry: ' + (endISO ? isoToDMY(endISO) : '?') + '\n\nThis updates the Google Sheet and generates the PDF.')) return;
    // loading button: takes ~10-15s (edit + settle delays + PDF export)
    var btn = $('#doBill');
    btn.disabled = true;
    btn.innerHTML = '<span class="mini-spinner"></span>Generating bill & PDF…';
    busy(true, 'Billing… generating PDF…');
    api('bill_customer', {
      customer: c.name,
      duration_months: m,
      amount: amt,
      description: $('#f_desc').value.trim(),
      payment_mode: $('#f_mode').value.trim(),
      address: $('#f_address').value.trim(),
      userid: $('#f_userid').value.trim(),
      invoice_date: invDate
    }).then(function (res) {
      busy(false);
      if (res.sheet_updated === false) {
        toast('PDF archived — but sheet update FAILED', true);
      } else {
        toast('Billed ✓ ' + res.invoice_no);
      }
      go(renderBillSuccess, { result: res });
      refreshData(false);
    }).catch(function (e) {
      busy(false);
      btn.disabled = false;
      btn.innerHTML = '🧾 Update Bill &amp; Advance Expiry';
      toast('Bill failed: ' + e.message, true);
    });
  };
}

function renderBillSuccess(app, params) {
  var r = params.result;
  var b64 = r.pdf_base64;
  var fname = r.pdf_name || (r.customer.replace(/\s+/g, '_') + '_' + r.invoice_no.replace(/[\\/]/g, '-') + '.pdf');
  app.innerHTML =
    '<div class="overlay"><div class="ohdr">' +
    '<button class="back" id="bk">←</button><h2>✅ Bill generated</h2></div>' +
    '<div class="wrap">' +
    '<div class="card" style="text-align:center;padding:22px">' +
    '<div style="font-size:44px">🧾</div>' +
    '<div style="font-size:21px;font-weight:800;margin:6px 0 2px">' + esc(r.invoice_no) + '</div>' +
    '<div class="meta-line">' + esc(r.customer) + ' · ' + esc(r.invoice_date) + '</div>' +
    '<div style="font-size:26px;font-weight:800;color:var(--accent);margin:10px 0">' + money(r.total) + '</div>' +
    '<div class="small">Basic ' + money(r.amount) + ' · CGST ' + money(r.cgst) + ' · SGST ' + money(r.sgst) + '</div>' +
    '</div>' +
    '<div class="card">' +
    '<div class="line" style="display:flex;justify-content:space-between;padding:4px 0"><span class="meta-line">New bill period</span><b>' + esc(r.period.start) + ' → ' + esc(r.period.end) + '</b></div>' +
    '<div class="line" style="display:flex;justify-content:space-between;padding:4px 0"><span class="meta-line">New expiry</span><b>' + esc(r.new_expire) + '</b></div>' +
    '<div class="line" style="display:flex;justify-content:space-between;padding:4px 0"><span class="meta-line">Duration</span><b>' + r.duration_months + ' months</b></div>' +
    '<div class="line" style="display:flex;justify-content:space-between;padding:4px 0"><span class="meta-line">Sheet update</span><b>' +
    (r.sheet_updated === false ? '<span style="color:var(--danger)">⚠ NOT updated</span>' : 'updated ✓') + '</b></div>' +
    (r.pdf_source ? '<div class="line" style="display:flex;justify-content:space-between;padding:4px 0"><span class="meta-line">PDF from</span><b>' +
      (String(r.pdf_source).indexOf('customer-doc') === 0 ? 'your Google Doc ✓' :
       String(r.pdf_source).indexOf('word-template') === 0 ? 'Word template ✓' : esc(r.pdf_source)) + '</b></div>' : '') +
    (r.pdf_settled === false ? '<div class="small" style="color:#b26a00;margin-top:8px">⚠ Google\'s PDF render cache was stale — the archived PDF will be replaced with the correct one within a few minutes automatically. Re-open it after ~3 min.</div>' : '') +
    (r.pdf_file ? '<div class="line" style="display:flex;justify-content:space-between;padding:4px 0"><span class="meta-line">Archived on Drive</span><b style="word-break:break-all;text-align:right">' + esc(r.pdf_file) + '</b></div>' : '') +
    '</div>' +
    (r.sheet_updated === false
      ? '<div class="card" style="background:var(--danger-soft);border-color:#fecaca"><div class="meta-line" style="color:var(--danger)"><b>⚠ PDF was generated and archived, but the sheet was NOT updated:</b><br>' + esc(r.sheet_error || 'unknown error') + '<br><br><b>Fix:</b> open this customer and bill again — the same invoice number will be reused and the sheet will be updated then.</div></div>'
      : '') +
    '<div class="btn-row"><button class="btn" id="viewPdf">👁 View</button>' +
    '<button class="btn" id="sharePdf">💬 WhatsApp / Share</button></div>' +
    '<div class="btn-row"><button class="btn primary" id="savePdf">⬇️ Download</button></div>' +
    '<div class="card"><div class="small">The bill was created from <b>your own Google Doc</b> (edited in place, no new docs) and the PDF is <b>archived on Drive</b> under BILLS/' + esc(r.customer) + '. "WhatsApp / Share" opens the Android share sheet — pick WhatsApp and the customer\'s chat. The phone keeps no copies.</div></div>' +
    '<div style="height:20px"></div></div></div>';
  $('#bk').onclick = back;
  $('#viewPdf').onclick = function () { openPdfViewer(r.invoice_no, b64, fname); };
  $('#savePdf').onclick = function () {
    downloadBlob(b64ToBlob(b64), fname);
    toast('Saved to Downloads ✓');
  };
  $('#sharePdf').onclick = function () { shareBlob(b64ToBlob(b64), fname); };
}

/* ------------------------------------------------------------- PDF VIEWER */
function openPdfViewer(invoiceNo, presetB64, presetName) {
  go(function (app, p) {
    app.innerHTML =
      '<div class="overlay"><div class="ohdr">' +
      '<button class="back" id="bk">←</button><h2>' + esc(p.invoiceNo) + '</h2></div>' +
      '<div id="pdfarea"><div class="wrap"><div class="empty"><div class="big">📄</div>Loading PDF…</div></div></div>' +
      '</div>';
    $('#bk').onclick = back;

    function show(b64) {
      var blob = b64ToBlob(b64);
      var name = p.presetName || (p.invoiceNo.replace(/[\\/]/g, '-') + '.pdf');
      $('#pdfarea').innerHTML =
        '<div class="wrap"><div class="pdf-pages" id="pages"></div></div>' +
        '<div class="pdf-actions">' +
        '<button class="btn" id="pshare">📤 Share</button>' +
        '<button class="btn primary" id="pdown">⬇️ Save to phone</button></div>';
      renderPdfPages_(blob, $('#pages'));
      $('#pdown').onclick = function () { downloadBlob(blob, name); toast('Saved to Downloads ✓'); };
      $('#pshare').onclick = function () { shareBlob(blob, name); };
    }

    if (p.presetB64) { show(p.presetB64); return; }

    // fetch the archived PDF from Drive (bills are archived on Drive only)
    if (!state.cfg) throw new Error('not connected');
    busy(true, 'Loading PDF from Drive…');
    api('get_pdf', { invoice_no: p.invoiceNo }).then(function (res) {
      busy(false);
      p.presetName = res.pdf_name;
      show(res.pdf_base64);
    }).catch(function (e) {
      busy(false);
      $('#pdfarea').innerHTML = '<div class="wrap"><div class="empty"><div class="big">⚠️</div>' +
        esc(e.message) + '<br><br><button class="btn" onclick="history.back()">Go back</button></div></div>';
    });
  }, { invoiceNo: invoiceNo, presetB64: presetB64 || null, presetName: presetName || null });
}

/** Render PDF pages onto canvases (mobile Chrome can't show PDFs in iframes).
 *  Falls back to the old iframe if pdf.js didn't load (e.g. no CDN access). */
function renderPdfPages_(blob, container) {
  if (!window.pdfjsLib) {
    container.innerHTML = '<iframe class="pdf-frame" src="' + URL.createObjectURL(blob) + '"></iframe>';
    return;
  }
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  var ab = blob.arrayBuffer ? blob.arrayBuffer() : new Promise(function (res, rej) {
    var fr = new FileReader();
    fr.onload = function () { res(fr.result); }; fr.onerror = rej;
    fr.readAsArrayBuffer(blob);
  });
  ab.then(function (buf) { return pdfjsLib.getDocument({ data: buf }).promise; })
    .then(function (pdf) {
      var n = Math.min(pdf.numPages, 6);           // bills are 1 page; cap anyway
      var next = function (i) {
        if (i > n) return Promise.resolve();
        return pdf.getPage(i).then(function (page) {
          var fit = (container.clientWidth || 320) / page.getViewport({ scale: 1 }).width;
          var vp = page.getViewport({ scale: Math.max(Math.min(fit, 2), 0.3) });
          var cv = document.createElement('canvas');
          cv.className = 'pdf-canvas';
          cv.width = vp.width; cv.height = vp.height;
          container.appendChild(cv);
          return page.render({ canvasContext: cv.getContext('2d'), viewport: vp })
            .promise.then(function () { return next(i + 1); });
        });
      };
      return next(1);
    })
    .catch(function () {
      container.innerHTML = '<iframe class="pdf-frame" src="' + URL.createObjectURL(blob) + '"></iframe>';
    });
}

/* -------------------------------------------------------------- SETTINGS */
function renderSettings(body) {
  var cs = state.data ? state.data.customers.length : 0;
  body.innerHTML = '<div class="wrap">' +
    '<div class="card"><div class="section-title" style="margin-top:0">Connection</div>' +
    '<div class="meta-line" style="word-break:break-all">' + esc(state.cfg ? state.cfg.url.replace(/\/exec.*$/, '/…') : '—') + '</div>' +
    '<div class="btn-row">' +
    '<button class="btn" id="s_test">Test connection</button>' +
    '<button class="btn" id="s_reconf">Change</button></div></div>' +

    '<div class="card"><div class="section-title" style="margin-top:0">Data</div>' +
    '<div class="meta-line">' + cs + ' customers loaded' + (state.savedAt ? ' · ' + new Date(state.savedAt).toLocaleString() : '') + '</div>' +
    '<div class="btn-row"><button class="btn" id="s_refresh">Refresh now</button></div></div>' +

    '<div class="card"><div class="section-title" style="margin-top:0">Install as app</div>' +
    '<div class="meta-line">Chrome menu (⋮) → <b>Add to Home screen</b>. It then opens fullscreen like a native app and works offline.</div>' +
    '<div class="btn-row"><button class="btn primary hidden" id="s_install">📲 Install</button></div></div>' +

    '<div class="card"><div class="section-title" style="margin-top:0">About</div>' +
    '<div class="meta-line">GST Bill Manager · bills are created by editing each customer\'s Google Doc and the PDF is archived on Drive under <b>BILLS/&lt;customer&gt;</b>. ' +
    'Desktop billing (run_billing.py) keeps working as before.</div></div>' +
    '<div style="height:20px"></div></div>';

  $('#s_test').onclick = function () {
    busy(true, 'Testing…');
    api('ping').then(function () { busy(false); toast('Server OK ✓'); })
      .catch(function (e) { busy(false); toast('Failed: ' + e.message, true); });
  };
  $('#s_reconf').onclick = function () { stack = [{ fn: renderSetup }]; renderTop(); };
  $('#s_refresh').onclick = function () { refreshData(true); };
  if (state.deferredInstall) {
    $('#s_install').className = 'btn primary';
    $('#s_install').onclick = function () {
      state.deferredInstall.prompt();
      state.deferredInstall = null;
      $('#s_install').className = 'btn primary hidden';
    };
  }
}

/* ------------------------------------------------------------------ boot */
window.addEventListener('beforeinstallprompt', function (e) {
  e.preventDefault();
  state.deferredInstall = e;
});
loadCfg();
loadCache();
if (!state.cfg) {
  stack = [{ fn: renderSetup }];
  renderTop();
} else {
  stack = [{ fn: renderTab, params: { tab: 'today' } }];
  renderTop();
  refreshData(false);
}
// refresh data when the app comes back to the foreground
document.addEventListener('visibilitychange', function () {
  if (!document.hidden && state.cfg && state.data && Date.now() - state.savedAt > 120000) refreshData(false);
});
