<div align="center">

# 💰 SpendLens

**Personal Budget Tracker for Google Sheets**

_Turn your bank CSV exports into a fully categorized, filterable budget tracker — entirely inside Google Sheets._

![Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-V8-4285F4?logo=google&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Google_Sheets-34A853?logo=google-sheets&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![No subscriptions](https://img.shields.io/badge/Cost-Free-brightgreen)

</div>

---

## Overview

SpendLens connects directly to your bank CSV exports and organizes your spending automatically. No third-party services, no subscriptions, no data leaving your Google account.

---

## Features

| | Feature | Description |
|---|---|---|
| 📂 | **Smart Import** | Upload CSV/XLS from CIBC, Tangerine or Amex Cobalt. Preview every transaction before importing. |
| 🔁 | **Duplicate Detection** | Already-imported transactions are automatically unchecked in the review screen. |
| 🏷️ | **Auto-categorization** | 120+ keyword rules map transactions to categories (e.g. `STARBUCKS` → Coffee Shop). |
| 🚫 | **Import Filters** | Sheet-managed rules to skip unwanted transactions — credit card payments, e-transfers, deposits. No code changes needed. |
| 💵 | **Budget Tracking** | Set monthly limits per category. Dashboard tracks your progress against each limit. |
| 📊 | **Dashboard** | Visual spending overview with charts and budget vs. actual comparison. |
| 📋 | **Category Rules** | Human-readable sheet showing all keyword rules, grouped by category. Fully editable. |
| 📤 | **Export Clean Template** | One-click export of a zero-data copy — all scripts and rules preserved, no personal transactions. |
| 📖 | **User Manual** | Built-in instruction sheet for non-technical users, accessible directly from the menu. |

---

## Supported Banks

| Bank | Format |
|---|---|
| **CIBC** | Chequing / Savings CSV export |
| **Tangerine** | Transaction history CSV export |
| **Amex Cobalt** | Statement XLS or CSV export |
| **Custom** | Register any CSV layout from the import modal (+ New Format) |

---

## Getting Started

### Manual Setup (no tooling required)

1. Open your Google Sheet (or create a new one)
2. Go to **Extensions → Apps Script**
3. Create one file for each `.js` and `.html` file in this repo, matching the filenames exactly
4. Paste the contents of each file
5. Open **Project Settings** and paste the contents of `appsscript.json` into the manifest editor
6. Save and reload the spreadsheet
7. A **💰 Budget Importer** menu will appear in the top bar

### Setup with clasp (recommended for developers)

```bash
# Install clasp globally
npm install -g @google/clasp

# Authenticate
clasp login

# Clone an existing script or push from local
clasp clone <your-script-id>
# or
clasp push
```

Create a `.clasp.json` in the project root (this file is gitignored — never commit it):

```json
{
  "scriptId": "YOUR_SCRIPT_ID_HERE",
  "rootDir": ""
}
```

---

## Project Structure

```
SpendLens_GoogleSheets/
├── Config.js          # Global constants, 120+ category rules, menu definition
├── Data.js            # Import logic, CSV parsers, category/filter/manual sheets
├── Engine.js          # Analysis engine, monthly spend aggregations
├── Dashboard.js       # Dashboard HTML builder
├── WebApp.js          # Budget sheet setup and web app entry point
├── Modal.html         # Import modal UI (HTML + CSS + JavaScript)
├── appsscript.json    # Apps Script project manifest
└── .gitignore         # Excludes .clasp.json and personal CSV files
```

---

## How It Works

### Import Flow

```
Bank website → Download CSV → Budget Importer menu
→ Select bank → Drop file → Review list → Confirm import
```

Duplicate transactions are auto-detected and unchecked by default.

### Categorization (two layers)

1. **Category Rules sheet** — your custom rules, editable directly in the sheet
2. **Config.js fallback** — 120+ built-in rules covering common Canadian merchants

Unmatched transactions are highlighted in **yellow** for manual review.

### Import Filters

The `Import Filters` sheet lets you define skip rules without touching code:

| Match Type | Example | Effect |
|---|---|---|
| `Amount → Positive` | — | Skips all deposits and credits |
| `Description` | `CREDIT CARD PAYMENT` | Skips any transaction containing that text |

Rules can be scoped to **All** banks or a specific bank (CIBC / Tangerine / Amex Cobalt).  
Toggle a rule on/off with the **Enabled** column — no deletion needed.  
Values are **case-insensitive** — `payment` and `PAYMENT` work the same way.

---

## Menu Reference

```
💰 Budget Importer
├── 📂 Import Bank File
├── ────────────────────
├── 🔄 Re-categorize uncategorized
├── ────────────────────
├── 📋 Category Rules
├── 🚫 Import Filters
├── 💵 Budget
├── 📖 User Manual
└── ────────────────────
    📤 Export Clean Template
```

All management sheets (Category Rules, Import Filters, Budget, Manual) follow the same toggle pattern: **first click opens**, **second click hides**.

---

## Sharing with Friends

Use **Budget Importer → 📤 Export Clean Template** to generate a shareable copy that contains:

- ✅ All category rules, import filters, CSV formats and scripts
- ❌ Zero personal transaction data

Share the link from the dialog. Your friend gets a fully working SpendLens from day one.

---

## Privacy

- All data stays inside your own Google account
- No external APIs, no third-party services, no tracking
- `.clasp.json` (script ID) and all CSV files are excluded from version control via `.gitignore`

---

## Requirements

- Google account with Google Sheets access
- _(For local development)_ Node.js + [clasp](https://github.com/google/clasp)

---

## License

[MIT](https://opensource.org/licenses/MIT) — free to use, modify and share.
