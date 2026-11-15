# GST Bill Manager — mobile app

A **phone app (installable PWA)** for your GST billing sheet. Everything runs on
the phone — no PC needed. The repo's existing files were not touched; this
folder is fully self-contained.

```
gst-mobile-app\
├── app\                  ← the phone app (PWA: HTML/CSS/JS, no build step)
├── appscript\
│   ├── Code.js           ← Apps Script backend = your original Code.js
│   │                        + mobile actions (list_all, bill_customer, make_pdf)
│   └── appsscript.json   ← same manifest as the sheet
└── test_backend.py       ← read-only tester (ping / list / pdf) from the PC
```

## What the app does

**Billing (customer-doc flow, v2):** each customer has ONE Google Doc (in the
`Gst_bills` folder on the telecom Drive). Billing **edits that doc in place**
— the previous bill's values are replaced with the new ones — then the PDF is
exported and **archived on Drive** under `BILLS/<Customer>/`. Only after the
PDF is safely archived does the app update the Google Sheet (advance expiry,
append the GST row). The phone keeps no copies: invoices are streamed from
Drive, viewable, shareable (**WhatsApp via the Android share sheet**) and
downloadable.

**Invoice history:** the customer page lists every invoice; tapping PDF
fetches the archived copy from Drive.

| Screen | What you see / do |
|---|---|
| **Home** | 🔴 bills **expiring today** (highlighted), overdue customers, next 7 days |
| **Expiries** | every customer sorted **nearest expiry first** — Expired / Today / 7 days / 30 days / Later |
| **Customers** | all GST-bill customers, searchable by name / user id / GSTIN |
| **Customer page** | expiry, duration, user id, GSTIN, address + **full invoice history** (amounts, payment status, PDF buttons) |
| **Bill / Update** | duration, amount, description, address, user id → one tap: next global invoice no → **PDF generated first** — the sheet is updated **only if the PDF succeeded** (a row added to `GST RECORDS` (Pending), expiry advanced, tab re-sorted) |
| **PDF** | every generated bill PDF is auto-saved **on the phone only** (viewable offline), saveable to Downloads, and shareable (WhatsApp etc.). Google Drive is never used |

The generated PDF replicates your Shreeganesh Telecom invoice layout
(company header, Buyer block, Invoice No/Date/User id/Support/PAN/GSTIN side
table, Bill Period, item table, CGST/SGST totals, amount in words, T&C and
bank details).

---

## One-time setup (≈10 minutes)

### 1. Deploy the new backend

1. Open your GST Google Sheet → **Extensions → Apps Script**.
2. **New project** (keeps the old automation untouched), name it
   e.g. `gst-mobile-backend`.
3. Delete the default code, paste the whole of `appscript\Code.js` (this folder).
4. Fill the two placeholders inside `setupSecrets()` with your `.env` values
   (`appscript_secret`, `sheet_id`), run `setupSecrets()` **once** (authorize
   when asked), then remove the pasted values again.
5. **Deploy → New deployment → Web app** — Execute as: **Me**, access:
   **Anyone** → Deploy → copy the `/exec` URL.

> The first billing from the phone will add two helper columns on its own:
> `Address` and `User id` (UPCOMING BILLS). Nothing else in the sheet changes,
> and the PC scripts keep working. Bill PDFs are generated in memory and sent
> straight to the phone — nothing is stored in Drive (the optional Word
> template below only needs Drive to read/delete its own throwaway copies).

### 2. Test from the PC (optional, read-only)

```powershell
python gst-mobile-app\test_backend.py ping
python gst-mobile-app\test_backend.py list --limit 10
python gst-mobile-app\test_backend.py pdf ST/053/2026-27   # any invoice with a PDF
```

### 3. Put the app on your phone (hosting)

The app is plain static files — any HTTPS host works. Easiest options:

**GitHub Pages (recommended — you already have this repo):**

The app files are already pushed to the **`gst-app`** branch (containing only
the app, at the branch root). Enable Pages once:

1. Open <https://github.com/rithikahir/llb-portal-automation/settings/pages>
2. Under **Build and deployment → Source** choose **Deploy from a branch**
3. Branch: **`gst-app`** · folder: **/(root)** → **Save**
4. Wait ~1 minute — your app URL becomes
   <https://rithikahir.github.io/llb-portal-automation/>

*(Any later app update: ask me to republish, or run the worktree commit +
`git push origin gst-app` again.)*

**Or Netlify Drop:** go to `https://app.netlify.com/drop` and drag the
`gst-mobile-app\app` folder in — done, gives you an HTTPS URL immediately.

### 4. Connect the phone

1. Open the hosted URL in Chrome on your phone.
2. First screen asks for the **Web App URL** (`/exec`) and the **shared
   secret** (the `appscript_secret` from `.env`). Connect — the dashboard loads.
3. Chrome menu (⋮) → **Add to Home screen** → it installs and opens
   fullscreen like a native app, works offline (cached UI + saved PDFs).

### 5. Bill layout — use the real Word template (one-time, optional but recommended)

By default the phone bills use a built-in HTML layout. To make every bill
**identical to the Word bills** (`HINDUSTAN_BILL.docx` layout), register the
bill template once:

1. The template file is ready: `GST_BILL_TEMPLATE.docx` in this folder —
   it is your real bill with `{{placeholders}}` baked in (same fonts, tables,
   layout). Regenerate it any time with `python make_template.py`.
2. Upload `GST_BILL_TEMPLATE.docx` to **Google Drive** (anywhere).
3. Double-click it → **Open with → Google Docs** (this converts it while
   keeping the layout). The URL now looks like
   `https://docs.google.com/document/d/`**`DOCUMENT_ID`**`/edit` — copy the
   long ID in the middle.
4. Register it (read-only for the sheet, stores just the ID):

   ```powershell
   python gst-mobile-app\test_backend.py template DOCUMENT_ID
   ```

5. One extra authorization: open the web app `/exec` URL once in a browser
   **as the account that deployed it** and accept the new
   *documents + drive* scopes when asked.
6. Test:

   ```powershell
   python gst-mobile-app\test_backend.py pdf ST/053/2026-27
   ```

   The line `[source: word-template]` means bills now render exactly like the
   Word documents. (`[source: html]` = no template registered yet.)

> How it works: for every bill the backend makes a **throwaway copy** of the
> template Google Doc, fills the placeholders, exports the PDF and **deletes
> the copy in the same execution** — nothing bill-related is ever kept in
> Drive, and the PDF is still created *before* the sheet is updated.

---

## Security notes

- The phone stores the URL + secret in the browser's local storage on your
  phone only. All traffic goes over HTTPS.
- Every API call is secret-gated. Bill PDFs are **never stored in Google
  Drive** — they exist in memory during generation and then live only on your
  phone. Old invoices can be regenerated at any time from the sheet with the
  `make_pdf` action (the app does this automatically when you tap a PDF that
  isn't stored on the phone).
## Desktop billing is untouched

`run_billing.py`, `auto_renew.py` etc. keep working exactly as before against
the original Apps Script project. The phone app and the PC script share the
same invoice numbering (global `ST/xxx/FY`), so you can bill from either side —
just not at the same moment.
