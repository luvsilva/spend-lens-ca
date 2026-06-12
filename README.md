# 💰 SpendLens — Personal Budget Tracker for Google Sheets

A Google Apps Script project that turns your bank CSV exports into a fully categorized, filterable budget tracker — all inside Google Sheets. No third-party services, no subscriptions, no data leaving your Google account.

---

## ✨ Features

| Feature | Description |
|---|---|
| 📂 **Smart Import** | Upload CSV files from CIBC, Tangerine or Amex Cobalt. Preview and deselect transactions before committing. |
| 🔁 **Duplicate Detection** | Already-imported transactions are automatically unchecked in the review screen. |
| 🏷️ **Auto-categorization** | 119+ keyword rules map transactions to categories (e.g. `STARBUCKS` → Coffee Shop). |
| 🚫 **Import Filters** | Sheet-managed rules to skip transactions you never want (credit card payments, e-transfers, deposits). No code changes needed. |
| 💵 **Budget Tracking** | Set monthly limits per category. The dashboard shows your progress against each limit. |
| 📊 **Dashboard** | Visual spending overview with charts and budget vs. actual comparison. |
| 📋 **Category Rules** | Grouped, human-readable sheet showing every keyword rule. Add new rules directly in the sheet. |
| 📤 **Export Clean Template** | One-click export of a zero-data copy to share with friends — all scripts and rules intact, no personal transactions. |
| 📖 **User Manual** | Built-in instruction sheet for non-technical users, accessible from the menu. |

---

## 🏦 Supported Banks

- **CIBC** — Chequing / Savings CSV export
- **Tangerine** — Transaction history CSV export
- **Amex Cobalt** — Statement XLS / CSV export
- **Custom formats** — Register any CSV layout directly from the import modal (+ New Format)

---

## 🚀 Getting Started

### 1 — Set up the Google Sheet

1. Open your Google Sheet (or create a new one)
2. Go to **Extensions → Apps Script**
3. Copy each `.js` and `.html` file from this repo into the Apps Script editor, matching the filenames
4. Copy `appsscript.json` content into the **Project Settings → appsscript.json** field
5. Save and reload the spreadsheet
6. A new **💰 Budget Importer** menu will appear in the top bar

### 2 — (Optional) Use clasp for local development

```bash
npm install -g @google/clasp
clasp login
clasp clone <your-script-id>   # or clasp push after editing locally
```

Create a `.clasp.json` in the project root:
```json
{
  "scriptId": "YOUR_SCRIPT_ID_HERE",
  "rootDir": ""
}
```

> `.clasp.json` is excluded from this repo via `.gitignore` — never commit your script ID.

---

## 📁 Project Structure

```
SpendLens_GoogleSheets/
├── Config.js          # Global constants, category rules, menu definition
├── Data.js            # Import logic, parsers, category/filter/manual sheets
├── Engine.js          # Analysis engine, monthly aggregations
├── Dashboard.js       # Dashboard HTML builder
├── WebApp.js          # Budget sheet setup, web app entry point
├── Modal.html         # Import modal UI (HTML/CSS/JS)
├── appsscript.json    # Apps Script manifest
└── .gitignore         # Excludes .clasp.json and personal data files
```

---

## 🗂️ How It Works

### Import Flow
1. Download a CSV from your bank
2. Open **Budget Importer → 📂 Import Bank File**
3. Select your bank, drop the file, review the transaction list
4. Uncheck anything you don't want, then click **Import**

### Categorization
Transactions are matched against keyword rules in two layers:
1. **Category Rules sheet** — your custom rules (editable in the sheet)
2. **Config.js fallback** — 119+ built-in rules covering common Canadian merchants

Unmatched transactions are highlighted in yellow for manual categorization.

### Import Filters
A dedicated sheet (`Import Filters`) lets you define skip rules:
- **Amount-based** — skip all positive amounts (deposits), negative amounts, etc.
- **Description-based** — skip any transaction containing a keyword (e.g. `CREDIT CARD PAYMENT`)
- Per-bank or global scope, toggled with a Yes/No column — no code changes needed

---

## 📤 Sharing with Friends

Use **Budget Importer → 📤 Export Clean Template** to create a copy with:
- ✅ All category rules, import filters, CSV formats, and scripts
- ❌ Zero personal transaction data

Share the generated link. Your friend gets a fully working SpendLens from day one.

---

## 🔒 Privacy

- All data stays in your own Google account
- No external APIs or third-party services
- `.clasp.json` (script ID) and any CSV files are excluded from version control via `.gitignore`

---

## 🛠️ Requirements

- Google account with Google Sheets access
- (For local development) Node.js + [clasp](https://github.com/google/clasp)

---

## 📄 License

MIT — free to use, modify and share.
