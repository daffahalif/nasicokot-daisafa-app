# Nasi Cokot Daisafa — Build Spec

Single source of truth for the implementation. `ABOUT.md` describes the product; this file describes the contract.

## Files

| File | Role |
|---|---|
| `Code.gs` | Apps Script backend: schema, API, queries, import |
| `index.html` | Entire frontend: HTML + CSS + JS + inline SVG icons |
| `appsscript.json` | Manifest: timezone `Asia/Jakarta`, web app access |

The script is **container-bound** to a Google Sheet converted from `NASI COKOT DAISAFA.xlsx`. Old tabs (`List Pesanan`, `Transaksi Baru`, `Pengeluaran`, `Dashboard Keuangan`, `Rasa`) stay as archive and are only read by `importLegacy()`.

No external libraries, CDNs, fonts, or images. Inline everything. All UI text is Indonesian. Money is integer Rupiah.

## Time and formats

- Timezone: `Asia/Jakarta` everywhere.
- `delivery_date` / `expense_date`: text `yyyy-MM-dd`.
- Timestamps: text `yyyy-MM-dd HH:mm:ss`.
- Money: integer, no decimals, no `"Rp"` in cells.
- IDs/phones: text format `@`.
- Enums are exact strings.

## Sheets

Header row is always row 1. No formulas in data cells. Freeze row 1.

### Customers
`customer_id | name | name_key | whatsapp | customer_key | aliases | merged_into | notes | active | created_at | updated_at`

- `customer_id`: `CUS-0001`
- `name_key`: lowercase, punctuation/emoji stripped, single spaces.
- `whatsapp`: digits only, `62…`, no `+`.
- `customer_key`: `wa:<phone>` if phone, else `name:<name_key>`.
- `aliases`: `|`-separated normalized variants.
- `merged_into`: tombstone target `customer_id`; merged rows are `active = FALSE`.
- Never hard-delete a customer. Merge = repoint orders, set `merged_into`, `active = FALSE`.

### Menu
`menu_id | name | short_name | price | cost | active | sort_order`

Seed:

| menu_id | name | short_name | price | cost | active | sort_order |
|---|---|---|---|---|---|---|
| M01 | Kulit Krengsengan | Kulit | 7000 | 6000 | TRUE | 1 |
| M02 | Mercon | Mercon | 7000 | 6000 | TRUE | 2 |
| M03 | Rica Kemangi | Rica | 7000 | 6000 | TRUE | 3 |
| M04 | Bumbu Rendang | Rendang | 7000 | 6000 | TRUE | 4 |
| M05 | Bumbu Bebek | Bebek | 7000 | 6000 | TRUE | 5 |

### Orders
`invoice_no | delivery_date | customer_id | customer_name | menu1_id | menu2_id | qty | unit_price | unit_cost | line_total | margin | payment_status | payment_method | order_status | notes | paid_at | created_at | updated_at`

- `invoice_no`: `NCD-YYMMDD-###`, PK, frozen at creation.
- `menu2_id` empty = single flavor. Mix = 1 portion at Rp7.000.
- `qty` = portions of the pair.
- `line_total = qty * unit_price`; `margin = line_total - qty * unit_cost`.
- `payment_status`: `Lunas` | `Belum Lunas`.
- `payment_method`: `Tunai` | `QRIS` | `Transfer` (blank unless Lunas).
- `order_status`: `Baru` | `Selesai` | `Batal`.
- `customer_name` is a snapshot for invoices.
- `unit_price` / `unit_cost` are frozen copies of the menu at write time.
- `Batal` is excluded from all reports, summaries, and receivables.

### Expenses
`expense_id | expense_date | category | description | amount | payment_method | notes | created_at | updated_at`

- `expense_id`: `EXP-0001`.
- `category`: `Beli Stok` | `Operasional` | `Lainnya`.
- `amount` > 0.

### Settings
`setting_key | setting_value | description | updated_at`

Seed keys:

```
business_name=Nasi Cokot Daisafa
owner_whatsapp=
qris_url=
invoice_footer=Terima kasih!
default_price=7000
default_cost=6000
invoice_prefix=NCD
timezone=Asia/Jakarta
currency=IDR
legacy_import_done=false
```

Invoice counters are stored as `invoice_seq:YYMMDD` = integer.

### Import_Log
`logged_at | source | source_row | level | issue | raw`

`level`: `INFO` | `WARN` | `SKIP`.

## Business rules

1. **Reseller model:** cost of goods is `unit_cost` (Rp6.000/porsi). Profit = `Σ margin`. `Expenses` records stock/operational cash out and is **not** subtracted from profit.
2. **Accrual:** profit counts all non-`Batal` orders by `delivery_date`, paid or not. Receivables = unpaid non-`Batal` `line_total`.
3. **Cash flow:** paid revenue minus expenses; shown separately.
4. **Mix:** two flavors on one portion. `menu1_id`/`menu2_id` sorted by `sort_order` at write.
5. **Invoice:** `NCD-` + `YYMMDD` of delivery date + `-` + 3-digit sequence per date. Generated under `LockService.getScriptLock()`.
6. **Schedule:** today/tomorrow are Jakarta dates; delivery-date based.

## API contract

`api(req)` where `req = { action, payload }`.

Success: `{ ok: true, data: ... }`
Failure: `{ ok: false, code: "VALIDATION"|"NOT_FOUND"|"SCHEMA"|"INTERNAL", message: "pesan Indonesia" }`

Return only JSON-safe values. Convert all dates/timestamps to strings.

### Actions

| Action | Payload | Data |
|---|---|---|
| `bootstrap` | `{}` | `{ today, tomorrow, settings, menu }` |
| `dashboard` | `{}` | `{ today:{date,portions,omzet,margin,orders}, month:{key,label,portions,omzet,margin}, receivables:{count,amount}, cash:{tunai,qris,transfer,total}, expenses_month, net_cash_month, unpaid:[{invoice_no,customer_name,whatsapp,amount,delivery_date}], qris_url }` |
| `orderBoard` | `{date}` | `{ date, date_label, summary_text, groups:[{type:"mix"\|"single",label,qty}], total, orders }` |
| `listOrders` | `{from?,to?,status?}` | `{ orders: [...] }` |
| `getOrder` | `{invoice_no}` | `{ order, customer, menu1, menu2, settings }` |
| `createOrder` | `{delivery_date,customer_id?,customer_name?,whatsapp?,menu1_id,menu2_id?,qty,payment_status,payment_method?,notes?}` | `{ order }` |
| `updateOrder` | `{invoice_no, ...order fields}` | `{ order }` |
| `cancelOrder` | `{invoice_no}` | `{ order }` |
| `listCustomers` | `{q?}` | `{ customers:[{...,order_count,last_order_date,total_spent}] }` |
| `saveCustomer` | `{customer_id?,name,whatsapp?,notes?,active?}` | `{ customer }` |
| `mergeCustomers` | `{source_id,target_id}` | `{ customer }` |
| `listMenu` | `{}` | `{ menu: [...] }` |
| `saveMenu` | `{menu_id,name,short_name,price,cost,active,sort_order}` | `{ menu }` |
| `listExpenses` | `{from?,to?}` | `{ expenses: [...] }` |
| `saveExpense` | `{expense_id?,expense_date,category,description,amount,payment_method?,notes?}` | `{ expense }` |
| `report` | `{period:"daily"\|"monthly",from?,to?,month?}` | `{ rows:[{key,label,portions,omzet,margin,expenses,net_cash}], totals }` |
| `getSettings` | `{}` | `{ settings }` |
| `saveSettings` | `{...keyValues}` | `{ settings }` |
| `importLegacy` | `{confirm:true}` | `{ already, counts:{customers,orders,expenses,duplicates,skipped}, warnings }` |

`ensureSchema()` is a **top-level function**, not an API action. Run it once from the Apps Script editor. It is non-destructive: create missing sheets, write headers on new sheets, seed Menu/Settings/Import_Log only when empty. Never clear, delete, or rewrite existing data. Never create a destructive `setupSheets()`.

### Order label

`menu1.short_name` if no `menu2`, else `short_name1 + "-" + short_name2` in `sort_order`.

### Summary text

Exact format:

```
Jumat, 23 Okt 2026
--- Menu Mix ---
3 - Kulit-Mercon
2 - Bebek-Mercon

--- Menu 1 Rasa ---
13 - Rendang
12 - Kulit
12 - Bebek
11 - Rica

Total: 53
```

- Day/month names in Indonesian: `Minggu Senin Selasa Rabu Kamis Jumat Sabtu`; `Jan Feb Mar Apr Mei Jun Jul Agu Sep Okt Nov Des`.
- Omit a section (and its blank line) when empty.
- Sort each section by qty desc, then label.
- Counts are portions, not rows.
- Exclude `Batal`.

## Import rules

Run `importLegacy()` once, after `ensureSchema()`. Process **Transaksi Baru first**, then **List Pesanan**, then **Pengeluaran**. Upsert customers, aggregate orders, assign invoice numbers chronologically. Log every warning/skip to `Import_Log`. Set `legacy_import_done=true` on success. If already true, return `{already:true}` without writing.

### Date parsing

Accept real `Date`, Excel serial number, or string. Excel serial: `new Date(Date.UTC(1899,11,30) + serial*86400000)`, format UTC parts. Strings: `yyyy-MM-dd`, `dd/MM/yyyy`, `dd-MM-yyyy`, `d MMMM yyyy`.

### Flavor normalization

`normalizeFlavor(s)`:
1. NFC normalize.
2. Remove `\u2060`, `\uFEFF`, `\u200B`-`\u200D`, `\uFE0F`, emoji.
3. Trim, collapse spaces, lowercase.
4. Strip leading `/^\s*\d+\s*[\.\)]\s*/`.
5. Map: `kulit krengsengan|kulit|krengsengan`→M01; `mercon`→M02; `rica|rica kemangi|kemangi`→M03; `bumbu rendang|rendang`→M04; `bumbu bebek|bebek`→M05.
6. Unknown → `null` + log, never guess.

### Transaksi Baru

Header row 1, data from row 2. Columns: `No, Hari, Tanggal, Nama Pemesan, Rasa 1, Rasa 2, Qty, Harga Jual, Harga Modal, Margin, Status Pesanan, Status Bayar, Metode Pembayaran`.

- Skip rows with no customer and no flavor.
- `qty` = col 7; if blank/0, derive `Harga Jual / 7000` if integer, else 1 + log.
- Recompute `line_total = qty * 7000`, `margin = qty * 1000`; ignore source totals.
- `Paid` → `Lunas`; `Unpaid`/blank → `Belum Lunas`.
- `CASH` → `Tunai`; `QRIS` → `QRIS`; `BSI`/`Transfer` → `Transfer`; blank + Lunas → blank + log.
- `order_status` = `Selesai` if delivery_date < today, else `Baru`.
- `paid_at` = delivery_date `12:00:00` when Lunas.
- Notes: `Impor: Transaksi Baru #<row>`.
- Aggregate by `(date, customer_id, menu1_id, menu2_id, payment_status, payment_method)`, sum qty.

### List Pesanan

Header row 5, data from row 6. Columns: `No, Nama, Rasa 1, Rasa 2, combined, Hari, Utk Tanggal, Grouping, Payment, Status`.

- **Forward-fill `Utk Tanggal`**: blank cells inherit the last non-empty date above.
- Each row = **1 portion** (source grand total 189 × 7.000 = 1.323.000).
- Flavor: prefer `Rasa 1`/`Rasa 2`; if both empty use `combined`, split on `-` after removing `\u2060`. If `Rasa 1` exists and `combined` conflicts, trust `Rasa 1` and log.
- Group by date block (forward-filled date). For each block:
  - `paidAmount` = sum of numeric `Payment` (col 9).
  - `paidQty = round(paidAmount / 7000)`.
  - Method from `Status` regex `/Paid\s*-\s*(BSI|QRIS|Cash)/i`: BSI→Transfer, QRIS→QRIS, Cash→Tunai.
  - Mark the first `min(paidQty, blockRows)` rows `Lunas`; the rest `Belum Lunas`. If `paidQty > blockRows`, mark all `Lunas` + log.
  - No method/status → `Belum Lunas`, blank method + log.
- Aggregate by `(date, customer_id, menu1_id, menu2_id, payment_status, payment_method)`, sum qty.
- Notes: `Impor: List Pesanan #<row>`.

### Dedupe

Key: `date|customer_key|menu1_id|menu2_id|qty|payment_status`. If the key already came from Transaksi Baru, skip the List Pesanan row and log `DUPLIKAT`. Only exact matches; log near-matches.

### Pengeluaran

Header row 1: A=`No`, B=`Hari`, C=`Tanggal`, D=`Keterangan`, E=`Nominal Pengeluaran`. Data from row 2. Skip rows with blank date or blank description+amount. Parse amount as integer (strip `Rp`, dots, spaces). `category = Beli Stok`, method blank, notes `Impor: Pengeluaran #<row>`.

### Customers

Upsert by `whatsapp` → `name_key` → `aliases`. Normalize phone to `62…`. `CUS-` + 4-digit sequence under lock. Empty name → `Tanpa Nama` + log. Do not auto-merge `Ko Billy`/`Koko Billy` or `Aga`/`Aga 2`; log them as candidates. `ibu`/`Ibu`, `mas fuad`/`Mas fuad`, `dimas`/`Dimas` dedupe naturally via `name_key`.

## UI requirements

- Single-page app in `index.html`. Sections shown/hidden in JS. No routing library.
- Bottom nav, 5 items: **Beranda, Pesanan, Pelanggan, Laporan, Lainnya**. `Lainnya` opens a sheet with **Menu, Nota, Keuangan, Pengaturan**.
- Pages and Indonesian labels:
  - Beranda: today/month profit + portions, today/tomorrow orders + copy summary, unpaid + WhatsApp links, QRIS.
  - Pesanan: create/edit/cancel order, date picker, customer picker/quick add, menu 1 + optional menu 2, qty, payment status/method, notes.
  - Pelanggan: search, add/edit, merge.
  - Menu: list/edit price, cost, active, sort order.
  - Nota: pick/search order, thermal receipt view, print + copy text.
  - Keuangan: cash summary + expenses list + add expense.
  - Laporan: daily/monthly table, period selector.
  - Pengaturan: business name, owner WhatsApp, QRIS URL, invoice footer, defaults, legacy import button with confirmation.
- Brand: `#1A3263` navy, `#547792` blue, `#EFD2B0` sand, `#FFC570` amber. Warm/cream background, navy primary, amber accent. Avoid generic Bootstrap look.
- Mobile-first at 390×844 (iPhone 12 Pro). `viewport-fit=cover`, safe-area bottom padding on nav, `height:100%` not `100vh`.
- Inputs `font-size:16px`; tap targets ≥44px; `touch-action:manipulation`.
- Thermal receipt: monospace, 80mm, `@media print` isolates `.receipt`, `@page { size: 80mm auto; margin: 0 }`. Always provide "Salin teks" fallback.
- Clipboard: try `navigator.clipboard.writeText`, fall back to hidden textarea + `document.execCommand('copy')` synchronously in the tap handler. Final fallback: modal with pre-selected text.
- WhatsApp: `https://wa.me/62…?text=…` as a real `<a target="_blank" rel="noopener">`.
- QRIS: if `qris_url` is set show the image; else show a styled placeholder "Atur gambar QRIS di Pengaturan".
- Frontend API wrapper: `call(action, payload)` returns a Promise around `google.script.run.withSuccessHandler().withFailureHandler()`. Add a 20s watchdog and a double-tap guard. Re-render from server data only.
- No `alert`/`confirm`; use in-page modals/toasts.
- All user-facing strings Indonesian.

## Safety

- `ensureSchema()` never deletes or clears. No `setupSheets()`.
- `importLegacy()` is idempotent via `legacy_import_done`.
- Writes use `LockService`; reads do not.
- Every write validates required fields and returns Indonesian error messages.
- After any column change, update `SCHEMA`, `COLUMN_MAP`-derived code, headers, and all readers/writers together.
