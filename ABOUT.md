## About this project

**Nasi Cokot Daisafa** is an app to manage nasi cokot's orders. It is built with Google Apps Script and uses Google Sheets as the database. this app refers to NASI COKOT DAISAFA.xlsx.

### The real app is only two files
- `Code.gs` is the backend.
- `index.html` is the whole front end (design and code in one file). Keep this name.
- Changes are pasted manually into the Apps Script editor. There is nothing to run locally.

### Files to ignore
- 

### What the app does
- Pages: Dashboard, Order, Customers, Menu, Invoice, Moneytary, Laporan (reports), Pengaturan (settings).
- use bottom navigation, uses in desktop and iphone 12 pro

### Users
- Me only

### Business terms
- Menu:
    1. kulit krengsengan
    2. ⁠mercon🌶️🌶️
    3. ⁠rica kemangi🌶️
    4. ⁠bumbu rendang
    5. ⁠bumbu bebek
- Order: Qty, Menu Ordered, Cust Name.
- Payment notes: PAID (QRIS/CASH), UNPAID.

### Data (Google Sheets)
Customers, Menu, Orders

### Design
- Use the brand colors: #1A3263, #547792, #EFD2B0. #FFC570.
- Invoices look like heat printer.
- Must work well on mobile, especially Iphone 12 Pro.

### DASHBOARD
1. monthly + daily profit, portion count
2. today order + summarizer (text to clipboard)
3. tomorrow order + summarizer (text to clipboard)
4. unpaid + customer whatsapp
5. QRIS

Summarized example:

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