// ============================================================
// Data.gs — REVIEW IMPORT v3 — preview + checkbox + category autocomplete
// ============================================================

// ── Import Modal ──────────────────────────────────────────────
function openReviewImportModalV5() {
  // IMPORTANT:
  // Do NOT call buildModalHTML() here because older files in Apps Script may also
  // define buildModalHTML() and override this one.
  // This unique function name guarantees the new Review Import modal opens.
  const html = HtmlService
    .createHtmlOutputFromFile('Modal')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(1120)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, "💰 Budget Importer — Review Before Import");
}

// Legacy wrapper: if your old menu still calls openImportModal, it will open the new review modal.
function openImportModal() {
  return openReviewImportModalV5();
}

function processFileImport(bank, fileContentBase64, fileName) {
  // Backward-compatible path: parse and import immediately.
  // The modal now uses previewFileImport() + commitReviewedImport(), but keeping this
  // function prevents old buttons/triggers from breaking.
  try {
    const preview = previewFileImport(bank, fileContentBase64, fileName);
    if (!preview.ok) return preview;
    const selected = preview.transactions.map(function(tx) {
      tx.include = true;
      return tx;
    });
    return commitReviewedImport(bank, selected);
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

/**
 * Step 1 — Parse the uploaded file and return a reviewable list.
 * Nothing is written to the spreadsheet here.
 * Used by the import modal to show preview + include checkbox + category dropdown.
 */
function previewFileImport(bank, fileContentBase64, fileName, customFormat) {
  try {
    let transactions = [];
    const filters = loadImportFilters();

    if (customFormat && customFormat.name) {
      if (!fileContentBase64) return { ok: false, error: "No file received." };
      const bytes   = Utilities.base64Decode(fileContentBase64);
      const blob    = Utilities.newBlob(bytes, "application/octet-stream", fileName || "import.csv");
      const csvText = blob.getDataAsString("UTF-8");
      const fmtCore = customFormat.name.replace(/[^\w\s]/g, '').trim().toUpperCase();
      if      (fmtCore === 'CIBC')       transactions = parseCIBC(csvText, filters);
      else if (fmtCore === 'TANGERINE')  transactions = parseTangerine(csvText, filters);
      else                               transactions = parseDynamicCsv_(csvText, customFormat);
    } else if (bank === "Amex" && String(fileContentBase64 || "").trim().startsWith("{")) {
      transactions = parseAmexParsedRows_(JSON.parse(fileContentBase64));
    } else {
      if (!fileContentBase64) return { ok: false, error: "No file received." };
      const bytes = Utilities.base64Decode(fileContentBase64);
      const blob = Utilities.newBlob(bytes, "application/octet-stream", fileName || "import.csv");
      const csvText = blobToCSV(bank, blob, fileName || "import.csv");

      if (bank === "CIBC") transactions = parseCIBC(csvText, filters);
      else if (bank === "Amex") transactions = parseAmex(csvText);
      else transactions = parseTangerine(csvText, filters);
    }

    if (!transactions || transactions.length === 0) {
      return { ok: false, error: "No valid transactions found." };
    }

    const smartRules = loadCategoryRules();
    const categories = getImportCategories_();
    const tz = Session.getScriptTimeZone();

    // Detect duplicates BEFORE showing the review screen.
    // Duplicate rows come unchecked by default and receive a Duplicate status.
    let existingSet = new Set();
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET);
      if (sheet) existingSet = loadExistingTransactions(sheet);
    } catch (e) {
      existingSet = new Set();
    }

    const reviewed = transactions.map(function(tx, idx) {
      const autoCat = autoCategoryHybrid(tx.desc, smartRules) || "";
      const duplicateKey = buildDuplicateKey(tx, tz);
      const isDuplicate = existingSet.has(duplicateKey);

      return {
        id: idx,
        include: !isDuplicate,
        isDuplicate: isDuplicate,
        duplicateKey: duplicateKey,
        dateISO: Utilities.formatDate(tx.date, tz, "yyyy-MM-dd"),
        dateLabel: Utilities.formatDate(tx.date, tz, "dd/MM/yyyy"),
        desc: tx.desc,
        amount: Math.round(parseAmountRobust(tx.amount) * 100) / 100,
        bank: tx.source || bank || "",
        category: autoCat,
        isUncategorized: !autoCat
      };
    });

    const sortedReviewed = sortReviewRowsForImport_(reviewed);

    return {
      ok: true,
      bank: bank,
      fileName: fileName || "",
      total: sortedReviewed.length,
      categories: categories,
      transactions: sortedReviewed
    };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}


/**
 * Keeps the preview screen focused on what matters:
 * 1) checked rows that will be imported
 * 2) unchecked non-duplicates
 * 3) duplicates at the bottom
 * Within each group, newest transactions appear first.
 */
function sortReviewRowsForImport_(rows) {
  return (rows || []).slice().sort(function(a, b) {
    const rankA = getReviewImportRank_(a);
    const rankB = getReviewImportRank_(b);
    if (rankA !== rankB) return rankA - rankB;

    const dateA = String(a.dateISO || "");
    const dateB = String(b.dateISO || "");
    if (dateA !== dateB) return dateB.localeCompare(dateA);

    return Number(b.amount || 0) - Number(a.amount || 0);
  });
}

function getReviewImportRank_(row) {
  if (row && row.include && !row.isDuplicate) return 0;
  if (row && !row.include && !row.isDuplicate) return 1;
  return 2;
}

/**
 * Step 2 — Commit only the rows checked by the user.
 * User-selected category is respected; blank category still imports and is highlighted yellow.
 */
function commitReviewedImport(bank, reviewedTransactions) {
  try {
    if (!Array.isArray(reviewedTransactions)) {
      return { ok: false, error: "Invalid reviewed transactions payload." };
    }

    const txs = [];
    reviewedTransactions.forEach(function(item) {
      if (!item || item.include === false) return;
      const date = parseReviewedDate_(item.dateISO || item.dateLabel);
      if (!isValidDate(date)) return;
      const desc = cleanDesc(item.desc || "");
      const amount = Math.abs(parseAmountRobust(item.amount));
      if (!desc || amount === 0) return;
      txs.push({
        date: date,
        desc: desc,
        amount: amount,
        source: item.bank || bank || "",
        category: String(item.category || "").trim()
      });
    });

    if (txs.length === 0) {
      return { ok: false, error: "No transactions selected to import." };
    }

    const result = writeTransactions(txs, bank);
    invalidateDashboardCache();

    const tz = Session.getScriptTimeZone();
    const skippedList = result.skippedTx.map(function(tx) {
      return Utilities.formatDate(tx.date, tz, "dd/MM/yyyy") + " | $" +
        tx.amount.toFixed(2) + " | " + tx.desc.substring(0, 55);
    });

    return {
      ok: true,
      added: result.added,
      skipped: result.skipped,
      skippedList: skippedList
    };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

function parseAmexParsedRows_(payload) {
  if (!payload || !payload.rows || payload.rows.length === 0) return [];
  const transactions = [];
  payload.rows.forEach(function(r) {
    const date = parseAmexDate(String(r.date || "").trim());
    if (!isValidDate(date)) return;
    const desc = cleanDesc(String(r.desc || "").trim());
    const amount = parseAmountRobust(r.amount);
    if (!desc || amount <= 0) return;
    transactions.push({ date: date, desc: desc, amount: amount, source: "Amex" });
  });
  return cancelReturns(transactions);
}

function parseReviewedDate_(value) {
  const s = String(value || "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const p = s.split("-");
    return new Date(+p[0], +p[1] - 1, +p[2]);
  }
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) {
    const p = s.split("/");
    return new Date(+p[2], +p[1] - 1, +p[0]);
  }
  return new Date(s);
}

function getImportCategories_() {
  const set = new Set();

  try {
    (CONFIG.CATEGORIES || []).forEach(function(rule) {
      if (rule && rule.category && !CONFIG.INCOME_CATEGORIES.includes(rule.category)) {
        set.add(rule.category);
      }
    });
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET);
    if (sheet && sheet.getLastRow() > 1) {
      const vals = sheet.getRange(2, CONFIG.COLS.CATEGORY, sheet.getLastRow() - 1, 1).getValues();
      vals.forEach(function(r) {
        const c = String(r[0] || "").trim();
        if (c && !CONFIG.INCOME_CATEGORIES.includes(c)) set.add(c);
      });
    }
  } catch (e) {}

  try {
    const budgets = getBudgets();
    Object.keys(budgets || {}).forEach(function(c) {
      if (c && !CONFIG.INCOME_CATEGORIES.includes(c)) set.add(c);
    });
  } catch (e) {}

  return Array.from(set).sort();
}

// Public helper for the modal if it needs to refresh categories later.
function getImportCategories() {
  return getImportCategories_();
}

function processAmexParsed(payload) {
  if (!payload || !payload.rows || payload.rows.length === 0)
    return { ok: false, error: "No valid transactions found in Amex file." };
  let transactions = [];
  payload.rows.forEach(r => {
    const date = parseAmexDate(String(r.date || "").trim());
    if (!isValidDate(date)) return;
    const desc   = cleanDesc(String(r.desc || "").trim());
    const amount = parseAmountRobust(r.amount);
    if (!desc || amount <= 0) return;
    transactions.push({ date, desc, amount, source: "Amex" });
  });
  transactions = cancelReturns(transactions);
  if (transactions.length === 0)
    return { ok: false, error: "No chargeable transactions found." };
  const result = writeTransactions(transactions, "Amex");
  invalidateDashboardCache();
  const tz = Session.getScriptTimeZone();
  const skippedList = result.skippedTx.map(tx =>
    Utilities.formatDate(tx.date, tz, "dd/MM/yyyy") + " | $" + tx.amount.toFixed(2) + " | " + tx.desc.substring(0, 55)
  );
  return { ok: true, added: result.added, skipped: result.skipped, skippedList };
}

// ── Parsers ───────────────────────────────────────────────────
function blobToCSV(bank, blob, fileName) {
  const lower = fileName.toLowerCase();
  if (bank === "Amex" || lower.endsWith(".xls") || lower.endsWith(".xlsx")) return parseAmexXlsBlob(blob);
  return blob.getDataAsString("UTF-8");
}

function parseAmexXlsBlob(blob) {
  const html = blob.getDataAsString("UTF-8");
  const trP = /<tr[^>]*>([\s\S]*?)<\/tr>/gi, tdP = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi, tagP = /<[^>]+>/g;
  const rows = []; let m;
  while ((m = trP.exec(html)) !== null) {
    const cells = []; let td; const re = new RegExp(tdP.source, "gi");
    while ((td = re.exec(m[1])) !== null) {
      let c = td[1].replace(tagP,"").replace(/&amp;/g,"&").replace(/&nbsp;/g," ")
        .replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&quot;/g,'"').replace(/&#39;/g,"'").replace(/\s+/g," ").trim();
      if (c.includes(",")) c = '"' + c.replace(/"/g,'""') + '"';
      cells.push(c);
    }
    if (cells.length) rows.push(cells.join(","));
  }
  if (!rows.length) throw new Error("Amex XLS: no table rows found — the file format may have changed.");
  return rows.join("\n");
}

function parseCIBC(csv, filters) {
  if (!filters) filters = loadImportFilters();
  const rows = parseCSV(csv), out = [];

  rows.forEach((row, i) => {
    if (!row || row.length < 3) return;
    if (i === 0 && String(row[0] || "").trim().toLowerCase().startsWith("date")) return;

    const date = parseCIBCDate(String(row[0] || "").trim());
    if (!isValidDate(date)) return;

    const desc   = String(row[1] || "").trim();
    const debit  = parseAmountRobust(row[2]);
    const credit = parseAmountRobust(row[3]);

    if (shouldSkipAmount_(credit, "CIBC", filters.amount))  return;
    if (shouldSkipDesc_(desc,    "CIBC", filters.description)) return;

    const amount = Math.abs(debit);
    if (!desc || amount === 0) return;

    out.push({ date, desc: cleanDesc(desc), amount, source: "CIBC" });
  });

  return cancelReturns(out);
}

function parseAmex(csv) {
  const rows = parseCSV(csv), out = []; let started = false;
  rows.forEach(row => {
    if (!started) { if (String(row[0]||"").trim() === "Date") started = true; return; }
    if (!row || row.length < 4) return;
    const date = parseAmexDate(String(row[0]||"").trim());
    if (!isValidDate(date)) return;
    const desc = String(row[2]||"").trim(), amount = parseAmountRobust(row[3]);
    if (!desc || amount === 0) return;
    out.push({ date, desc: cleanDesc(desc), amount, source: "Amex" });
  });
  return cancelReturns(out);
}

function parseTangerine(csv, filters) {
  if (!filters) filters = loadImportFilters();
  const rows = parseCSV(csv);
  const out = [];

  rows.forEach((row, i) => {
    if (!row || row.length < 5) return;
    if (i === 0 && String(row[0] || "").toLowerCase().trim() === "date") return;

    const date = parseTangerineDate(String(row[0] || "").trim());
    if (!isValidDate(date)) return;

    const transactionType = String(row[1] || "").trim();
    const name      = String(row[2] || "").trim();
    const memo      = String(row[3] || "").trim();
    const rawAmount = parseAmountRobust(row[4]);

    if (rawAmount === 0) return;
    if (shouldSkipAmount_(rawAmount, "Tangerine", filters.amount)) return;

    const combined = name + " " + memo + " " + transactionType;
    if (shouldSkipDesc_(combined, "Tangerine", filters.description)) return;

    let desc = [name, memo]
      .filter(Boolean)
      .join(" - ")
      .replace(/^Visa Debit\s*-\s*Purchase\s*-\s*/i, "")
      .replace(/^Interac\s*-\s*Purchase\s*-\s*/i, "")
      .replace(/^INTERAC\s+e-Transfer\s+(From|To):\s*/i, "e-Transfer: ")
      .replace(/^Automated Banking Machine\s*-?\s*/i, "ATM WITHDRAWAL ")
      .replace(/^ABM\s*-?\s*/i, "ATM WITHDRAWAL ");

    desc = cleanDesc(desc);
    if (!desc) return;

    out.push({ date, desc, amount: Math.abs(rawAmount), source: "Tangerine" });
  });

  return cancelReturns(out);
}

function cancelReturns(transactions) {
  const groups = {};
  const tz = Session.getScriptTimeZone();
  transactions.forEach((tx, idx) => {
    const day = Utilities.formatDate(tx.date, tz, "yyyy-MM-dd");
    const key = day + "|" + tx.desc.toUpperCase().substring(0, 30);
    if (!groups[key]) groups[key] = [];
    groups[key].push({ idx, amount: tx.amount });
  });
  const toRemove = new Set();
  Object.values(groups).forEach(entries => {
    if (entries.length < 2) return;
    const pos = entries.filter(e => e.amount > 0).sort((a,b) => b.amount - a.amount);
    const neg = entries.filter(e => e.amount < 0).sort((a,b) => a.amount - b.amount);
    neg.forEach(n => {
      const match = pos.find(p => Math.round(p.amount*100) === Math.round(Math.abs(n.amount)*100));
      if (match) { toRemove.add(n.idx); toRemove.add(match.idx); pos.splice(pos.indexOf(match),1); }
    });
  });
  return transactions.filter((_, idx) => !toRemove.has(idx));
}

// ── Write & Sort ──────────────────────────────────────────────
function writeTransactions(transactions, bank) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET);
  if (!sheet) throw new Error(`Sheet "${CONFIG.TARGET_SHEET}" not found.`);
  const existing   = loadExistingTransactions(sheet);
  const tz         = Session.getScriptTimeZone();
  const smartRules = loadCategoryRules();
  const newRows = [], uncatIdx = [], skippedTx = [];
  let added = 0, skipped = 0;

  transactions.forEach(tx => {
    const key = buildDuplicateKey(tx, tz);
    if (existing.has(key)) { skipped++; skippedTx.push(tx); return; }
    const category = String(tx.category || "").trim() || autoCategoryHybrid(tx.desc, smartRules);
    const row = [
      new Date(tx.date.getFullYear(), tx.date.getMonth(), tx.date.getDate(), tx.date.getHours(), tx.date.getMinutes(), tx.date.getSeconds()),
      new Date(tx.date.getFullYear(), tx.date.getMonth(), 1),
      tx.desc, category, tx.amount, tx.source || bank || ""
    ];
    if (!category) uncatIdx.push(newRows.length);
    newRows.push(row); existing.add(key); added++;
  });

  if (newRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, 6).setValues(newRows);
    sheet.getRange(startRow, 1, newRows.length, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    sheet.getRange(startRow, 2, newRows.length, 1).setNumberFormat("dd/MM/yyyy");
    sheet.getRange(startRow, 5, newRows.length, 1).setNumberFormat("\"$\"#,##0.00");
    uncatIdx.forEach(i => sheet.getRange(startRow+i, 1, 1, 6).setBackground(CONFIG.UNCATEGORIZED_COLOR));
  }
  sortSheetByDate(sheet);
  return { added, skipped, skippedTx };
}

function sortSheetByDate(sheet) {
  const lastRow = sheet.getLastRow(), lastCol = Math.max(sheet.getLastColumn(), 6);
  if (lastRow < 2) return;
  const firstCell = sheet.getRange(1, 1).getValue();
  const hasHeader = typeof firstCell === "string" && firstCell.trim() !== "";
  const startRow = hasHeader ? 2 : 1, rowCount = lastRow - startRow + 1;
  if (rowCount < 2) return;
  sheet.getRange(startRow, 1, rowCount, lastCol).sort({ column: CONFIG.COLS.TIMESTAMP, ascending: true });
}

// ── Re-categorize ─────────────────────────────────────────────
function recategorizeUncategorized() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET);
  if (!sheet || sheet.getLastRow() < 2) { SpreadsheetApp.getUi().alert("No data found."); return; }
  const smartRules  = loadCategoryRules();
  const lastRow     = sheet.getLastRow();
  const data        = sheet.getRange(2, 1, lastRow-1, 6).getValues();
  const backgrounds = sheet.getRange(2, 1, lastRow-1, 1).getBackgrounds();
  let fixed = 0, cleared = 0;
  data.forEach((row, i) => {
    const cat     = String(row[CONFIG.COLS.CATEGORY-1]||"").trim();
    const bg      = backgrounds[i][0];
    const isYellow = bg === CONFIG.UNCATEGORIZED_COLOR || bg === "#fff9c4";
    if (isYellow && cat) {
      sheet.getRange(i+2, 1, 1, 6).setBackground(null); cleared++;
    } else if (isYellow && !cat) {
      const desc = String(row[CONFIG.COLS.DESCRIPTION-1]||"").trim();
      const newCat = autoCategoryHybrid(desc, smartRules);
      if (newCat) { sheet.getRange(i+2, CONFIG.COLS.CATEGORY).setValue(newCat); sheet.getRange(i+2, 1, 1, 6).setBackground(null); fixed++; }
    }
  });
  invalidateDashboardCache();
  const msg = [];
  if (fixed > 0)   msg.push("✅ Auto-categorized: " + fixed);
  if (cleared > 0) msg.push("🧹 Yellow cleared: " + cleared);
  if (!msg.length) msg.push("✅ Nothing to update.");
  SpreadsheetApp.getUi().alert(msg.join("\n"));
}

/**
 * Reclassifica uma transação diretamente da planilha.
 * Chamado pelo drill-down do dashboard.
 * Recebe: { date, desc, amount, bank, newCategory }
 * Localiza a linha pelo trio (date, desc, amount) e atualiza a categoria.
 */
function reclassifyTransaction(txInfo) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET);
    if (!sheet) return { ok: false, error: "Sheet not found." };
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: false, error: "No data." };

    const tz   = Session.getScriptTimeZone();
    const data = sheet.getRange(2, 1, lastRow-1, 6).getValues();
    const targetAmt = Math.round(parseAmountRobust(txInfo.amount) * 100);

    for (let i = 0; i < data.length; i++) {
      const row  = data[i];
      const rawB = row[1];
      const desc = String(row[2]||"").trim();
      const amt  = Math.round(parseAmountRobust(row[4]) * 100);
      const bank = String(row[5]||"").trim();

      if (desc !== txInfo.desc) continue;
      if (amt  !== targetAmt)  continue;
      if (txInfo.bank && bank !== txInfo.bank) continue;

      // Verifica data (coluna B = Expense Date = 1º do mês)
      let rowDate = null;
      if (rawB instanceof Date && isValidDate(rawB)) rowDate = rawB;
      else if (typeof rawB === "number" && rawB > 1000)
        rowDate = new Date(Date.UTC(1899,11,30) + rawB*86400000);
      if (!rowDate) continue;
      const rowMk = Utilities.formatDate(rowDate, tz, "yyyy-MM");
      if (rowMk !== txInfo.monthKey) continue;

      // Encontrou — atualiza categoria e limpa amarelo
      sheet.getRange(i+2, CONFIG.COLS.CATEGORY).setValue(txInfo.newCategory);
      sheet.getRange(i+2, 1, 1, 6).setBackground(null);
      invalidateDashboardCache();
      return { ok: true };
    }
    return { ok: false, error: "Transaction not found in spreadsheet." };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Summary ───────────────────────────────────────────────────
function showSummary() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET);
  if (!sheet || sheet.getLastRow() < 2) { SpreadsheetApp.getUi().alert("No data found."); return; }
  const data = sheet.getRange(2, CONFIG.COLS.CATEGORY, sheet.getLastRow()-1, 2).getValues();
  const catMap = {}; let uncategorized = 0;
  data.forEach(row => {
    const cat = String(row[0]||"").trim();
    if (!cat) { uncategorized++; return; }
    catMap[cat] = (catMap[cat]||0) + parseAmountRobust(row[1]);
  });
  const sorted = Object.entries(catMap).sort((a,b) => b[1]-a[1]);
  const total  = sorted.reduce((s,[,v]) => s+v, 0);
  let msg = "Total: $" + total.toFixed(2) + "\n\n";
  sorted.slice(0,15).forEach(([cat,amt]) => { msg += cat + ": $" + amt.toFixed(2) + "\n"; });
  if (uncategorized > 0) msg += "\n⚠️ " + uncategorized + " uncategorized";
  SpreadsheetApp.getUi().alert(msg);
}

// ── Categorização híbrida ─────────────────────────────────────
function loadCategoryRules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rs = ss.getSheetByName(CONFIG.RULES_SHEET);
  if (!rs || rs.getLastRow() < 2) return [];
  const data = rs.getRange(2, 1, rs.getLastRow()-1, 2).getValues();
  const rules = [];
  const skipPrefixes = ['─','—','➕','ℹ','#','=','—'];
  data.forEach(function(row) {
    const keyword  = String(row[0]||"").trim();
    const category = String(row[1]||"").trim();
    if (!keyword || !category) return;
    if (skipPrefixes.some(function(p) { return keyword.startsWith(p); })) return;
    rules.push({ keyword: keyword.toUpperCase(), category: category, priority: 1 });
  });
  rules.sort(function(a, b) { return b.keyword.length - a.keyword.length; });
  return rules;
}

function autoCategoryHybrid(desc, smartRules) {
  const upper = String(desc||"").toUpperCase();
  if (smartRules && smartRules.length > 0)
    for (const rule of smartRules)
      if (upper.includes(rule.keyword)) return rule.category;
  return autoCategory(desc);
}

function autoCategory(desc) {
  const upper = String(desc||"").toUpperCase();
  for (const rule of CONFIG.CATEGORIES)
    for (const kw of rule.match)
      if (upper.includes(kw.toUpperCase())) return rule.category;
  return "";
}

// ── CategoryRules Sheet ───────────────────────────────────────
function setupCategoryRulesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let rs = ss.getSheetByName(CONFIG.RULES_SHEET);
  if (rs) { rs.clearContents(); rs.clearFormats(); }
  else     { rs = ss.insertSheet(CONFIG.RULES_SHEET); }

  // Row 1 — header
  rs.getRange(1,1,1,2)
    .setValues([["Keyword  —  matches text in your transaction description","Category  —  label applied in the spreadsheet"]])
    .setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold").setFontSize(11);
  rs.setRowHeight(1, 32);

  // Row 2 — info banner
  rs.getRange(2,1,1,2).merge()
    .setValue("ℹ️   To hide this sheet: Budget Importer menu  →  📋 Category Rules")
    .setBackground("#e0f2fe").setFontColor("#0369a1").setFontStyle("italic")
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  rs.setRowHeight(2, 26);
  rs.setFrozenRows(2);

  // Group CONFIG.CATEGORIES by category, preserving declaration order
  const grouped = {};
  const order = [];
  (CONFIG.CATEGORIES || []).forEach(function(rule) {
    if (!grouped[rule.category]) { grouped[rule.category] = []; order.push(rule.category); }
    (rule.match || []).forEach(function(kw) { grouped[rule.category].push(kw); });
  });

  const EMOJIS = {
    "Groceries":"🛒","Restaurant":"🍽","Fast Food":"🍔","Coffee Shop":"☕",
    "GAS":"⛽","Streamming":"📺","GYM":"🏋","Supplements":"💊",
    "Mobile":"📱","Internet":"🌐","BCHydro":"💡","Car Payment":"🚗",
    "Appartament Rent":"🏠","Fees":"💸","Banking Fees":"🏦","ICBC":"🛡",
    "Car Insurance":"🛡","Parking":"🅿","Public Transportation":"🚌",
    "Private Transportation":"🚕","Car Maintenance":"🔧","Clothes":"👕",
    "SkinCare":"✨","Pharmacy":"💊","Health":"🏥","Dental":"🦷",
    "House Maintanence":"🏡","Dog Food":"🐕","Vacation":"✈",
    "Leisure":"🎭","Night Club":"🎉","BarberShop":"✂",
    "Tax Return":"📊","Rent Insurance":"🏠","Courses":"📚",
    "Driver School":"🚗","BCLiquor":"🍷","Transfer":"💸","Cash Withdrawal":"💵"
  };

  // Alternating row colors per group for visual separation
  const ROW_COLORS = ["#ffffff","#f8fafc"];
  const SECTION_BG = "#e2e8f0";
  const SECTION_FG = "#334155";

  let row = 3;
  order.forEach(function(cat, gi) {
    const keywords = grouped[cat];
    if (!keywords || keywords.length === 0) return;
    const emoji    = EMOJIS[cat] || "📌";
    const rowColor = ROW_COLORS[gi % 2];

    // Section header (merged across both columns)
    rs.getRange(row, 1, 1, 2).merge()
      .setValue(emoji + "   " + cat.toUpperCase())
      .setBackground(SECTION_BG).setFontColor(SECTION_FG)
      .setFontWeight("bold").setFontSize(10).setVerticalAlignment("middle");
    rs.setRowHeight(row, 26);
    row++;

    // Keyword rows — batch write for performance
    const vals = keywords.map(function(kw) { return [kw, cat]; });
    rs.getRange(row, 1, vals.length, 2).setValues(vals)
      .setBackground(rowColor).setFontSize(11);
    row += vals.length;
  });

  rs.setColumnWidth(1, 300);
  rs.setColumnWidth(2, 190);
  if (rs.isSheetHidden()) rs.showSheet();
  ss.setActiveSheet(rs);
}

function toggleCategoryRulesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rs = ss.getSheetByName(CONFIG.RULES_SHEET);
  if (rs && !rs.isSheetHidden()) { rs.hideSheet(); return; }
  setupCategoryRulesSheet();
}

function extractMeaningfulTokens(desc) {
  const cleaned = desc
    .replace(/^(Point of Sale\s*-\s*Interac\s*RETAIL PURCHASE\s*\d*\s*)/i,"")
    .replace(/^(Electronic Funds Transfer\s*(PREAUTHORIZED DEBIT\s*\d*\s*)?)/i,"")
    .replace(/^(Branch Transaction\s*)/i,"")
    .replace(/^(Internet Banking\s*(INTERNET TRANSFER|E-TRANSFER)\s*\d*\s*)/i,"")
    .replace(/\d{6,}/g,"").replace(/[#*@]/g," ").replace(/\s{2,}/g," ").trim().toUpperCase();
  const stopWords = new Set([
    "THE","AND","FOR","FROM","WITH","RETAIL","PURCHASE","DEBIT","PREAUTHORIZED",
    "ELECTRONIC","TRANSFER","FUNDS","INTERAC","POINT","SALE","BRANCH","TRANSACTION",
    "INTERNET","BANKING","DEPOSIT","PAYMENT","CANADA","VANCOUVER","BURNABY","SURREY",
    "RICHMOND","NORTH","SOUTH","EAST","WEST","ONLINE","STORE","MARKET","INC","LTD",
    "CORP","LLC","TORONTO","ONTARIO","BRITISH","COLUMBIA","WWW","COM","SERVICES",
    "PERSONAL","HOUSEHOLD","EXPENSES","PROFESSIONAL","FINANCIAL","RETAIL","GROCERY",
    "HOTEL","ENTERTAINMENT","RECREATION","HEALTH","EDUCATION",
    "AMAZON","AMZN","AMAZONCA","AMAZON.CA","MKTP"
  ]);
  const words  = cleaned.split(/\s+/);
  const tokens = new Set();
  words.forEach(w => { if (w.length>=4 && !stopWords.has(w) && !/^\d+$/.test(w)) tokens.add(w); });
  for (let i=0; i<words.length-1; i++) {
    const w1=words[i], w2=words[i+1];
    if (w1.length>=3 && w2.length>=3 && !stopWords.has(w1) && !stopWords.has(w2)) tokens.add(w1+" "+w2);
  }
  return [...tokens];
}

// ── Helpers ───────────────────────────────────────────────────
function parseCIBCDate(str) {
  const s = String(str||"").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) { const p=s.split("-"); return new Date(+p[0],+p[1]-1,+p[2]); }
  return new Date(s);
}
function parseAmexDate(str) {
  const M={Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
  const s=String(str||"").replace(/\./g,"").trim().split(/\s+/);
  if (s.length>=3) return new Date(+s[2],M[s[1]]??0,+s[0]);
  return new Date(str);
}
function parseTangerineDate(str) {
  const s=String(str||"").trim();
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) { const p=s.split("/"); return new Date(+p[2],+p[0]-1,+p[1]); }
  return new Date(s);
}
function isValidDate(d)  { return d instanceof Date && !isNaN(d.getTime()); }
function cleanDesc(desc) { return String(desc||"").replace(/\s{2,}/g," ").replace(/\t/g," ").trim().substring(0,120); }
function parseAmountRobust(v) {
  if (!v && v!==0) return 0;
  if (typeof v==="number") return v;
  let s=String(v).trim().replace(/\$/g,"").replace(/\s/g,"");
  if (/^-?\d{1,3}(\.\d{3})*,\d+$/.test(s)) s=s.replace(/\./g,"").replace(",",".");
  else s=s.replace(/,/g,"");
  return Number(s.replace(/[^0-9.-]/g,""))||0;
}
function parseCSV(text) {
  const norm=String(text||"").replace(/\r\n/g,"\n").replace(/\r/g,"\n");
  const rows=[]; let row=[],cur="",inQ=false;
  for (let i=0;i<norm.length;i++) {
    const ch=norm[i],nx=norm[i+1];
    if (ch==='"'){if(inQ&&nx==='"'){cur+='"';i++;}else inQ=!inQ;}
    else if (ch===","&&!inQ){row.push(cur);cur="";}
    else if (ch==="\n"&&!inQ){row.push(cur);if(row.some(c=>String(c).trim()))rows.push(row.map(c=>String(c).trim()));row=[];cur="";}
    else cur+=ch;
  }
  row.push(cur);
  if (row.some(c=>String(c).trim())) rows.push(row.map(c=>String(c).trim()));
  return rows;
}
function loadExistingTransactions(sheet) {
  const set=new Set(), lastRow=sheet.getLastRow();
  if (lastRow<2) return set;
  const tz=Session.getScriptTimeZone();
  const data=sheet.getRange(2,1,lastRow-1,5).getValues();
  data.forEach(row => {
    const dateVal=row[0], desc=row[2], amount=row[4];
    if (!dateVal||!desc) return;
    set.add(buildDuplicateKey({date:dateVal,desc,amount},tz));
  });
  return set;
}
function buildDuplicateKey(tx, tz) {
  const timezone = tz || Session.getScriptTimeZone();
  const d = tx.date instanceof Date && isValidDate(tx.date)
    ? Utilities.formatDate(tx.date, timezone, "yyyy-MM-dd")
    : String(tx.date||"").trim();
  return `${d}|${String(tx.desc||"").toLowerCase().trim().substring(0,40)}|${Math.round(parseAmountRobust(tx.amount)*100)}`;
}

// ── ImportFilters — Exclusion Rules Registry ─────────────────
const IMPORT_FILTERS_SHEET = "ImportFilters";

const IMPORT_FILTERS_SEEDS_ = [
  ["Skip Deposits",             "Amount",       "Positive",                "All",        "Yes"],
  ["Credit Card Payment",       "Description",  "PAYMENT - THANK YOU",     "All",        "Yes"],
  ["Pre-Auth Payment",          "Description",  "PRE-AUTHORIZED PAYMENT",  "All",        "Yes"],
  ["Preauth Payment",           "Description",  "PREAUTHORIZED PAYMENT",   "All",        "Yes"],
  ["Preauth Debit",             "Description",  "PREAUTHORIZED DEBIT",     "All",        "Yes"],
  ["Credit Card Payment Long",  "Description",  "CREDIT CARD PAYMENT",     "All",        "Yes"],
  ["Credit Card Pymt",          "Description",  "CREDIT CARD PYMT",        "All",        "Yes"],
  ["CC Payment",                "Description",  "CC PAYMENT",              "All",        "Yes"],
  ["Visa Payment",              "Description",  "VISA PAYMENT",            "All",        "Yes"],
  ["Visa Bill",                 "Description",  "VISA BILL",               "All",        "Yes"],
  ["Mastercard Payment",        "Description",  "MASTERCARD PAYMENT",      "All",        "Yes"],
  ["Master Card Payment",       "Description",  "MASTER CARD PAYMENT",     "All",        "Yes"],
  ["Internet Transfer",         "Description",  "INTERNET TRANSFER",       "All",        "Yes"],
  ["E-Transfer",                "Description",  "E-TRANSFER",              "All",        "Yes"],
  ["Amex Bill",                 "Description",  "AMEX BILL",               "Tangerine",  "Yes"],
  ["American Express Payment",  "Description",  "AMERICAN EXPRESS",        "CIBC",       "Yes"],
  ["CIBC Card Products",        "Description",  "CIBC CARD PRODUCTS",      "CIBC",       "Yes"],
  ["CIBC Credit Card",          "Description",  "CIBC CREDIT CARD",        "CIBC",       "Yes"],
  ["Transferred",               "Description",  "TRANSFERRED",             "Tangerine",  "Yes"],
];

function _getImportFilterBankOptions_() {
  const options = ["All"];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fmtSh = ss.getSheetByName(CSV_FORMATS_SHEET);
  if (fmtSh && fmtSh.getLastRow() > 1) {
    fmtSh.getRange(2, 1, fmtSh.getLastRow() - 1, 1).getValues().forEach(function(row) {
      const clean = String(row[0] || "").replace(/[^\w\s]/g, "").trim();
      if (clean && options.indexOf(clean) < 0) options.push(clean);
    });
  }
  if (options.length === 1) { options.push("CIBC", "Tangerine", "Amex Cobalt"); }
  return options;
}

function _refreshImportFiltersBankDropdown_(sh) {
  const opts = _getImportFilterBankOptions_();
  sh.getRange(3, 4, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(opts, true).build());
}

function _getImportFiltersSheet_(create) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(IMPORT_FILTERS_SHEET);
  if (!sh && create) {
    sh = ss.insertSheet(IMPORT_FILTERS_SHEET);

    // Row 1 — header
    sh.getRange(1,1,1,5)
      .setValues([["Filter Name","Match Type","Value  (case-insensitive)","Bank","Enabled"]])
      .setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold").setFontSize(11);
    sh.setRowHeight(1, 32);

    // Row 2 — info banner
    sh.getRange(2,1,1,5).merge()
      .setValue("ℹ️   To hide this sheet: Budget Importer menu  →  🚫 Import Filters")
      .setBackground("#e0f2fe").setFontColor("#0369a1").setFontStyle("italic")
      .setHorizontalAlignment("left").setVerticalAlignment("middle");
    sh.setRowHeight(2, 26);
    sh.setFrozenRows(2);

    // Data starts at row 3
    sh.getRange(3,1,IMPORT_FILTERS_SEEDS_.length,5).setValues(IMPORT_FILTERS_SEEDS_);
    IMPORT_FILTERS_SEEDS_.forEach(function(row, i) {
      sh.getRange(i+3,1,1,5).setBackground(row[1] === "Amount" ? "#fef3c7" : "#f0f9ff");
    });

    sh.getRange(3,2,1000,1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(["Amount","Description"],true).build());
    _refreshImportFiltersBankDropdown_(sh);
    sh.getRange(3,5,1000,1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(["Yes","No"],true).build());

    sh.setColumnWidth(1,200); sh.setColumnWidth(2,130);
    sh.setColumnWidth(3,270); sh.setColumnWidth(4,130); sh.setColumnWidth(5,80);
    sh.hideSheet();
  }
  return sh;
}

function loadImportFilters() {
  const sh = _getImportFiltersSheet_(false);
  if (!sh || sh.getLastRow() < 2) return { description: [], amount: [] };
  const data = sh.getRange(2,1,sh.getLastRow()-1,5).getValues();
  const descRules = [], amountRules = [];
  data.forEach(function(row) {
    const type    = String(row[1]||"").trim().toUpperCase();
    const value   = String(row[2]||"").trim().toUpperCase();
    const bank    = String(row[3]||"").trim().toUpperCase();
    const enabled = String(row[4]||"").trim().toUpperCase();
    if (!type || !value || enabled !== "YES") return;
    if (type === "DESCRIPTION") descRules.push({ bank, value });
    else if (type === "AMOUNT")  amountRules.push({ bank, direction: value });
  });
  return { description: descRules, amount: amountRules };
}

function shouldSkipDesc_(text, bank, rules) {
  const upper = String(text||"").toUpperCase();
  const bankU = bank.toUpperCase();
  for (var i = 0; i < rules.length; i++) {
    if (rules[i].bank !== "ALL" && rules[i].bank !== bankU) continue;
    if (upper.includes(rules[i].value)) return true;
  }
  return false;
}

function shouldSkipAmount_(amount, bank, rules) {
  const bankU = bank.toUpperCase();
  for (var i = 0; i < rules.length; i++) {
    if (rules[i].bank !== "ALL" && rules[i].bank !== bankU) continue;
    if (rules[i].direction === "POSITIVE" && amount > 0) return true;
    if (rules[i].direction === "NEGATIVE" && amount < 0) return true;
  }
  return false;
}

function toggleImportFiltersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(IMPORT_FILTERS_SHEET);
  if (!sh) { _getImportFiltersSheet_(true); sh = ss.getSheetByName(IMPORT_FILTERS_SHEET); }
  if (!sh) return;
  if (!sh.isSheetHidden()) { sh.hideSheet(); return; }
  _refreshImportFiltersBankDropdown_(sh);
  sh.showSheet();
  ss.setActiveSheet(sh);
}



// ── User Manual Sheet ─────────────────────────────────────────
function setupManualSheet() {
  const MANUAL_SHEET = "Manual";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MANUAL_SHEET);
  if (sh) { sh.clearContents(); sh.clearFormats(); }
  else sh = ss.insertSheet(MANUAL_SHEET);

  const sections = [
    { title: "📂   IMPORTING TRANSACTIONS", rows: [
      ["Step 1 — Download your bank file",    "Log in to your bank website and download the transaction history as a CSV file (or .xls for Amex Cobalt)."],
      ["Step 2 — Open the importer",          "Click  Budget Importer menu  →  📂 Import Bank File."],
      ["Step 3 — Select your bank",           "Click your bank button: CIBC, Tangerine or Amex Cobalt. If your bank is not listed, click + New Format to register a custom layout."],
      ["Step 4 — Drop your file",             "Drag and drop the downloaded file onto the area, or click it to browse your computer."],
      ["Step 5 — Review transactions",        "A list appears. Uncheck any rows you don't want to import. Already-imported transactions are automatically unchecked."],
      ["Step 6 — Confirm",                    "Click  Import X transactions. They will appear in the Expenses_Form sheet."],
    ]},
    { title: "🏷️   CATEGORIES", rows: [
      ["How categories work",                 "Every transaction is automatically matched to a category based on its description (e.g. STARBUCKS → Coffee Shop)."],
      ["Yellow highlight = needs attention",  "A yellow row means the transaction could not be categorized. Click the Category cell and type the correct category manually."],
      ["View all rules",                      "Budget Importer menu  →  📋 Category Rules. Shows every keyword grouped by category."],
      ["Add a new rule",                      "In the Category Rules sheet, add a new row: keyword in column A (e.g. MY STORE), category in column B (e.g. Groceries)."],
      ["Re-apply rules to old rows",          "After adding new rules, go to  Budget Importer  →  🔄 Re-categorize uncategorized  to apply them to existing yellow rows."],
    ]},
    { title: "🚫   IMPORT FILTERS", rows: [
      ["What filters do",                     "Automatically skip transactions you never want imported — like credit card payments, e-transfers between accounts, or deposits."],
      ["Open filters",                        "Budget Importer menu  →  🚫 Import Filters."],
      ["Match Type: Description",             "Skips any transaction whose description contains the Value text. Not case-sensitive — you can type in any capitalization."],
      ["Match Type: Amount — Positive",       "Skips deposits and credits (money coming in to your account)."],
      ["Bank column",                         "All = applies to every bank.  CIBC / Tangerine / Amex Cobalt = applies only to that bank."],
      ["Enabled = No",                        "Temporarily disables a rule without deleting it. Change back to Yes to re-enable it."],
      ["Add a new filter",                    "Scroll to the bottom of the sheet and add a new row. Fill in all 5 columns: Filter Name, Match Type, Value, Bank, Enabled."],
    ]},
    { title: "💵   BUDGET", rows: [
      ["How budgets work",                    "Set a monthly spending limit per category. The Dashboard shows your progress against each limit."],
      ["Open / reset the Budget sheet",       "Budget Importer menu  →  💵 Budget.  YES = rebuild with your latest categories.  NO = just open it as-is."],
      ["Set a limit",                         "Type the dollar amount in the  Monthly Budget ($)  column. Example: type 500 for Groceries."],
      ["No limit",                            "Leave 0 for categories you don't want to track against a budget."],
      ["Notes column",                        "Optional — add a personal reminder (e.g. 'includes gym + supplements')."],
    ]},
    { title: "📤   EXPORT CLEAN TEMPLATE", rows: [
      ["What it creates",                     "A full copy of this spreadsheet with all your transactions removed. Category rules, filters and scripts are kept intact."],
      ["How to export",                       "Budget Importer menu  →  📤 Export Clean Template  →  click Yes  →  copy the link from the dialog."],
      ["Share with a friend",                 "Send the copied link. Your friend gets a fully working SpendLens with zero personal data from you."],
      ["Your friend's first steps",           "Open the link → accept the copy → go to Budget Importer menu → 📂 Import Bank File and start importing their own transactions."],
    ]},
  ];

  // Row 1 — title
  sh.getRange(1,1,1,2).merge()
    .setValue("💰   SpendLens — User Manual")
    .setBackground("#0f172a").setFontColor("#f8fafc")
    .setFontWeight("bold").setFontSize(14).setVerticalAlignment("middle").setHorizontalAlignment("left");
  sh.setRowHeight(1, 40);

  // Row 2 — info banner
  sh.getRange(2,1,1,2).merge()
    .setValue("ℹ️   To hide this sheet: Budget Importer menu  →  📖 User Manual")
    .setBackground("#e0f2fe").setFontColor("#0369a1").setFontStyle("italic")
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  sh.setRowHeight(2, 26);
  sh.setFrozenRows(2);

  const SECTION_BG = "#334155";
  const ROW_COLORS = ["#ffffff", "#f8fafc"];
  let row = 3;

  sections.forEach(function(sec, si) {
    sh.getRange(row,1,1,2).merge()
      .setValue(sec.title)
      .setBackground(SECTION_BG).setFontColor("#ffffff")
      .setFontWeight("bold").setFontSize(11).setVerticalAlignment("middle");
    sh.setRowHeight(row, 28);
    row++;
    const rowColor = ROW_COLORS[si % 2];
    sec.rows.forEach(function(r) {
      sh.getRange(row,1).setValue(r[0]).setFontWeight("bold").setFontSize(11)
        .setBackground(rowColor).setVerticalAlignment("top").setWrap(true);
      sh.getRange(row,2).setValue(r[1]).setFontSize(11)
        .setBackground(rowColor).setVerticalAlignment("top").setWrap(true);
      sh.setRowHeight(row, 42);
      row++;
    });
  });

  sh.setColumnWidth(1, 250);
  sh.setColumnWidth(2, 540);
  if (sh.isSheetHidden()) sh.showSheet();
  ss.setActiveSheet(sh);
}

function toggleManualSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Manual");
  if (sh && !sh.isSheetHidden()) { sh.hideSheet(); return; }
  setupManualSheet();
}

// ── Export Clean Template ─────────────────────────────────────
function exportCleanTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "Export Clean Template",
    "This will create a copy of this spreadsheet with all transaction data removed.\n\n" +
    "KEPT: category rules, import filters, CSV formats, all scripts.\n" +
    "CLEARED: all transactions (Expenses_Form) and Analysis sheet.\n\nContinue?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  const tz   = Session.getScriptTimeZone();
  const name = "SpendLens - Clean Template - " + Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const file = DriveApp.getFileById(ss.getId()).makeCopy(name);
  const copy = SpreadsheetApp.openById(file.getId());
  const expSheet = copy.getSheetByName(CONFIG.TARGET_SHEET);
  if (expSheet && expSheet.getLastRow() > 1)
    expSheet.getRange(2, 1, expSheet.getLastRow()-1, expSheet.getLastColumn()).clearContent().setBackground(null);
  ["Analysis"].forEach(function(n) {
    const s = copy.getSheetByName(n);
    if (s && s.getLastRow() > 1) s.getRange(2,1,s.getLastRow()-1,s.getLastColumn()).clearContent();
  });
  const url = "https://docs.google.com/spreadsheets/d/" + file.getId();
  ui.alert("Template Created!", "Clean template ready — zero personal data.\n\n" + name + "\n\n" + url, ui.ButtonSet.OK);
}

// ── CsvFormats — Dynamic Import Registry ──────────────────────
const CSV_FORMATS_SHEET = "CsvFormats";

function _getCsvFormatsSheet_(create) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CSV_FORMATS_SHEET);
  if (!sh && create) {
    sh = ss.insertSheet(CSV_FORMATS_SHEET);
    sh.getRange(1,1,1,7).setValues([["FormatName","DateCol","DescCol","AmountCol","DateFormat","AmountSign","SkipRows"]]);
    sh.getRange(1,1,1,7).setBackground("#4f46e5").setFontColor("#fff").setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.setColumnWidth(1,180);
    // Seed built-in formats so they appear alongside custom ones
    sh.getRange(2,1,2,7).setValues([
      ["🏦 CIBC",      0, 1, 2, "yyyy-MM-dd", "positive", 1],
      ["🍊 Tangerine", 0, 2, 4, "MM/dd/yyyy", "negative", 1]
    ]);
    sh.hideSheet();
  }
  return sh;
}

function seedBuiltinCsvFormats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CSV_FORMATS_SHEET);
  if (!sh) { _getCsvFormatsSheet_(true); return; }
  const builtins = [
    ["🏦 CIBC",      0, 1, 2, "yyyy-MM-dd", "positive", 1],
    ["🍊 Tangerine", 0, 2, 4, "MM/dd/yyyy", "negative", 1]
  ];
  const existing = sh.getLastRow() >= 2
    ? sh.getRange(2,1,sh.getLastRow()-1,1).getValues().map(function(r){ return String(r[0]||"").trim(); })
    : [];
  builtins.forEach(function(row) {
    if (!existing.includes(row[0])) sh.appendRow(row);
  });
}

function getCsvFormats() {
  try {
    const sh = _getCsvFormatsSheet_(false);
    if (!sh || sh.getLastRow() < 2) return { ok: true, formats: [] };
    const data = sh.getRange(2, 1, sh.getLastRow()-1, 7).getValues();
    const formats = [];
    data.forEach(function(row) {
      const name = String(row[0]||"").trim();
      if (!name) return;
      formats.push({
        name:       name,
        dateCol:    Number(row[1]||0),
        descCol:    Number(row[2]||1),
        amountCol:  Number(row[3]||2),
        dateFormat: String(row[4]||"auto").trim(),
        amountSign: String(row[5]||"positive").trim(),
        skipRows:   Number(row[6]||1)
      });
    });
    return { ok: true, formats: formats };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function saveCsvFormat(config) {
  try {
    const name = String(config.name||"").trim();
    if (!name) return { ok: false, error: "Format name is required." };
    const sh = _getCsvFormatsSheet_(true);
    const row = [
      name, Number(config.dateCol), Number(config.descCol),
      Number(config.amountCol), String(config.dateFormat||"auto"),
      String(config.amountSign||"positive"), Number(config.skipRows||1)
    ];
    if (sh.getLastRow() >= 2) {
      const data = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]||"").trim() === name) {
          sh.getRange(i+2,1,1,7).setValues([row]);
          return { ok: true, updated: true };
        }
      }
    }
    sh.appendRow(row);
    return { ok: true, updated: false };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function hideBuiltinBank(name) {
  try {
    const n = String(name||"").trim();
    if (!n) return { ok: false, error: "Name required." };
    const props = PropertiesService.getUserProperties();
    const hidden = JSON.parse(props.getProperty('hiddenBuiltins') || '[]');
    if (!hidden.includes(n)) { hidden.push(n); props.setProperty('hiddenBuiltins', JSON.stringify(hidden)); }
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function getCombinedFormatsData() {
  try {
    seedBuiltinCsvFormats();
    const props = PropertiesService.getUserProperties();
    const hidden = JSON.parse(props.getProperty('hiddenBuiltins') || '[]');
    const customRes = getCsvFormats();
    return { ok: true, formats: customRes.formats || [], hiddenBuiltins: hidden };
  } catch(e) {
    return { ok: true, formats: [], hiddenBuiltins: [] };
  }
}

function deleteCsvFormat(name) {
  try {
    const n = String(name||"").trim();
    if (!n) return { ok: false, error: "Name required." };
    const sh = _getCsvFormatsSheet_(false);
    if (!sh || sh.getLastRow() < 2) return { ok: true, deleted: false };
    const data = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]||"").trim() === n) {
        sh.deleteRow(i+2);
        return { ok: true, deleted: true };
      }
    }
    return { ok: true, deleted: false };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function parseDynamicCsv_(csvText, fmt) {
  const rows = parseCSV(csvText);
  const out  = [];
  const skip = Number(fmt.skipRows||1);
  rows.forEach(function(row, i) {
    if (i < skip) return;
    const dateStr = String(row[fmt.dateCol]||"").trim();
    const desc    = String(row[fmt.descCol]||"").trim();
    const rawAmt  = parseAmountRobust(row[fmt.amountCol]);
    if (!dateStr || !desc || rawAmt === 0) return;
    const date = parseDynamicDate_(dateStr, fmt.dateFormat);
    if (!isValidDate(date)) return;
    let amount;
    if (fmt.amountSign === "negative") {
      if (rawAmt >= 0) return;
      amount = Math.abs(rawAmt);
    } else {
      if (rawAmt <= 0) return;
      amount = rawAmt;
    }
    out.push({ date: date, desc: cleanDesc(desc), amount: amount, source: String(fmt.name||"Custom") });
  });
  return cancelReturns(out);
}

function parseDynamicDate_(str, format) {
  const s = String(str||"").trim();
  if (!s) return new Date("invalid");
  if (!format || format === "auto") {
    if (/^\d{4}-\d{2}-\d{2}$/.test(s))                           return parseCIBCDate(s);
    if (/^\d{1,2}\s+[A-Za-z]{3}\.?\s+\d{4}$/.test(s))           return parseAmexDate(s);
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s))                    return parseTangerineDate(s);
    return new Date(s);
  }
  if (format === "dd MMM yyyy" || format === "dd Mon yyyy")       return parseAmexDate(s);
  if (format === "yyyy-MM-dd")                                    return parseCIBCDate(s);
  if (format === "MM/dd/yyyy" || format === "M/d/yyyy")          return parseTangerineDate(s);
  if (format === "dd/MM/yyyy") {
    const p = s.split("/");
    if (p.length === 3) return new Date(+p[2], +p[1]-1, +p[0]);
  }
  return new Date(s);
}

// ── Modal HTML ────────────────────────────────────────────────
function buildReviewImportModalHTML_() {
  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<script async src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"><\/script>
<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Google Sans',Arial,sans-serif;background:#f8fafc;padding:18px;color:#0f172a;font-size:13px;}
h2{font-size:18px;font-weight:700;color:#4338ca;margin-bottom:6px;}
.subtitle{font-size:12px;color:#64748b;margin-bottom:14px;line-height:1.4;}
.banks{display:flex;gap:6px;margin-bottom:10px;flex-wrap:wrap;align-items:center;}
.bank-btn{flex:1;min-width:72px;padding:9px 8px;border:2px solid #e2e8f0;border-radius:12px;background:#fff;cursor:pointer;text-align:center;transition:all .15s;box-shadow:0 1px 2px rgba(15,23,42,.04);font-family:inherit;font-size:13px;display:inline-flex;align-items:center;justify-content:center;gap:4px;}
.bank-btn:hover{border-color:#6366f1;background:#eef2ff;}
.bank-btn.active{border-color:#4f46e5;background:#eef2ff;color:#4338ca;}
.bank-btn .name{font-weight:700;font-size:12px;}
.bank-btn.custom-fmt{border-color:#c7d2fe;}
.bank-btn.custom-fmt.active{border-color:#4f46e5;background:#e0e7ff;}
.bank-btn.new-fmt{border-color:#86efac;background:#f0fdf4;color:#166534;flex:none;}
.bank-btn.new-fmt:hover{border-color:#4ade80;background:#dcfce7;}
.del-fmt{display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;border-radius:50%;background:rgba(239,68,68,.15);color:#ef4444;font-size:11px;font-weight:900;margin-left:4px;flex-shrink:0;line-height:1;cursor:pointer;}
.del-fmt:hover{background:#ef4444;color:#fff;}
.name-input-row{display:none;align-items:center;gap:8px;margin-bottom:12px;background:#f0fdf4;border:1px solid #86efac;border-radius:10px;padding:10px 14px;flex-wrap:wrap;}
.name-input-row label{font-size:12px;font-weight:700;color:#166534;white-space:nowrap;}
.name-input-row input{flex:1;min-width:180px;max-width:280px;padding:8px 12px;border:1px solid #86efac;border-radius:8px;font-size:13px;outline:none;background:#fff;}
.name-input-row input:focus{border-color:#22c55e;box-shadow:0 0 0 3px rgba(34,197,94,.1);}
.dropzone{border:2px dashed #cbd5e1;border-radius:16px;background:#fff;padding:28px 20px;text-align:center;cursor:pointer;transition:all .15s;margin-bottom:12px;position:relative;box-shadow:0 8px 24px rgba(15,23,42,.05);}
.dropzone.drag{border-color:#4f46e5;background:#eef2ff;transform:scale(1.005);}
.dropzone:hover{border-color:#4f46e5;background:#eef2ff;}
.dropzone input[type=file]{display:none;}
.file-selected{display:none;align-items:center;gap:10px;background:#f0fdf4;border:1px solid #86efac;border-radius:10px;padding:10px 14px;margin-bottom:12px;}
.fs-name{font-weight:600;color:#166534;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.fs-remove{cursor:pointer;color:#64748b;font-size:16px;border:none;background:none;}
.actions{display:flex;align-items:center;gap:10px;margin-top:12px;flex-wrap:wrap;}
.btn{padding:10px 16px;border:none;border-radius:10px;font-size:13px;font-weight:700;cursor:pointer;transition:all .15s;}
.btn-primary{background:#4f46e5;color:#fff;} .btn-primary:hover{background:#4338ca;}
.btn-secondary{background:#e2e8f0;color:#0f172a;} .btn-secondary:hover{background:#cbd5e1;}
.btn-danger{background:#fee2e2;color:#991b1b;}
.btn-success{background:#dcfce7;color:#166534;border:1px solid #86efac;} .btn-success:hover{background:#bbf7d0;}
.btn:disabled{background:#94a3b8;color:#fff;cursor:not-allowed;}
#status{font-size:13px;flex:1;min-width:200px;}
.ok{color:#16a34a;font-weight:700;}.err{color:#ef4444;font-weight:700;}.inf{color:#4f46e5;font-weight:700;}.warn{color:#d97706;font-weight:700;}
/* ── Setup Panel ── */
.setup-panel{background:#fff;border:1px solid #c7d2fe;border-radius:16px;padding:16px;margin-bottom:12px;box-shadow:0 4px 16px rgba(79,70,229,.08);}
.setup-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;}
.setup-title{font-size:15px;font-weight:800;color:#4338ca;}
.setup-row{margin-bottom:12px;}
.setup-label{font-size:11px;font-weight:700;color:#475569;display:block;margin-bottom:4px;text-transform:uppercase;letter-spacing:.04em;}
.setup-input{width:100%;padding:9px 12px;border:1px solid #cbd5e1;border-radius:10px;font-size:13px;background:#fff;outline:none;}
.setup-input:focus{border-color:#4f46e5;box-shadow:0 0 0 3px rgba(79,70,229,.1);}
.setup-dropzone{border:2px dashed #c7d2fe;border-radius:12px;background:#eef2ff;padding:18px;text-align:center;cursor:pointer;position:relative;margin-bottom:12px;}
.setup-dropzone input[type=file]{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%;}
.setup-dropzone:hover{border-color:#4f46e5;background:#e0e7ff;}
/* ── Mapping UI ── */
.mapping-wrap{display:flex;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden;margin-bottom:14px;}
.map-col-left{flex:0 0 148px;background:#f8fafc;border-right:1px solid #e2e8f0;}
.map-col-arrow{flex:0 0 42px;background:#f1f5f9;border-right:1px solid #e2e8f0;display:flex;flex-direction:column;}
.map-col-right{flex:1;background:#fff;}
.map-col-hdr{padding:8px 12px;background:#4f46e5;color:#fff;font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.05em;min-height:34px;display:flex;align-items:center;}
.map-field-row{display:flex;align-items:center;padding:10px 12px;border-bottom:1px solid #f1f5f9;min-height:60px;}
.map-field-row:last-child{border-bottom:none;}
.map-arrow-row{display:flex;align-items:center;justify-content:center;min-height:60px;border-bottom:1px solid #f1f5f9;font-size:18px;color:#4f46e5;font-weight:900;}
.map-arrow-row:last-child{border-bottom:none;}
.map-field-inner{display:flex;flex-direction:column;gap:2px;}
.map-field-name{font-weight:700;font-size:12px;color:#1e293b;}
.map-field-req{color:#ef4444;font-size:10px;}
.map-right-inner{display:flex;flex-direction:column;gap:3px;width:100%;}
.map-select{width:100%;padding:7px 9px;border:1px solid #cbd5e1;border-radius:8px;font-size:12px;background:#fff;outline:none;}
.map-select:focus{border-color:#4f46e5;}
.map-sample{font-size:10px;color:#94a3b8;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.setup-opts{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px;}
.setup-opt label{font-size:11px;font-weight:700;color:#475569;display:block;margin-bottom:4px;text-transform:uppercase;letter-spacing:.04em;}
.setup-opt select{width:100%;padding:8px 10px;border:1px solid #cbd5e1;border-radius:8px;font-size:12px;background:#fff;outline:none;}
.setup-opt select:focus{border-color:#4f46e5;}
.sample-wrap{background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:10px;margin-bottom:12px;overflow-x:auto;}
.sample-lbl{font-size:11px;font-weight:700;color:#64748b;margin-bottom:6px;}
.sample-wrap table{width:100%;border-collapse:collapse;font-size:11px;}
.sample-wrap th{background:#e2e8f0;padding:5px 8px;text-align:left;font-weight:700;color:#334155;}
.sample-wrap td{padding:5px 8px;border-bottom:1px solid #f1f5f9;color:#334155;}
/* ── Review Panel ── */
.review-panel{display:none;background:#fff;border:1px solid #e2e8f0;border-radius:16px;margin-top:14px;box-shadow:0 10px 30px rgba(15,23,42,.07);overflow:hidden;}
.review-header{padding:12px 14px;background:#f8fafc;border-bottom:1px solid #e2e8f0;display:flex;gap:10px;align-items:center;justify-content:space-between;flex-wrap:wrap;}
.review-title{font-weight:800;color:#1e293b;}
.review-tools{display:flex;gap:8px;align-items:center;flex-wrap:wrap;}
.search{padding:8px 10px;border:1px solid #cbd5e1;border-radius:10px;min-width:180px;font-size:12px;}
.tiny{font-size:11px;color:#64748b;}
.table-wrap{max-height:300px;overflow:auto;}
table{width:100%;border-collapse:collapse;font-size:12px;}
th{position:sticky;top:0;background:#f1f5f9;color:#334155;text-align:left;padding:8px;border-bottom:1px solid #cbd5e1;z-index:2;}
td{padding:7px 8px;border-bottom:1px solid #eef2f7;vertical-align:middle;}
tr:hover{background:#f8fafc;}
tr.excluded{opacity:.50;background:#f8fafc;} tr.duplicate{opacity:.62;background:#f1f5f9;} tr.uncat{background:#fff7ed;}
.amount{text-align:right;font-variant-numeric:tabular-nums;font-weight:700;}
.desc{max-width:260px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.cat-cell{position:relative;min-width:230px;}
.cat-input{width:215px;padding:8px 10px;border:1px solid #cbd5e1;border-radius:10px;background:#fff;font-size:12px;outline:none;transition:box-shadow .12s,border-color .12s;}
.cat-input:focus{border-color:#4f46e5;box-shadow:0 0 0 3px rgba(79,70,229,.12);}
.cat-input.missing{border-color:#f97316;background:#fff7ed;}
.autocomplete-menu{position:absolute;left:8px;top:42px;width:250px;max-height:220px;overflow:auto;background:#fff;border:1px solid #cbd5e1;border-radius:12px;box-shadow:0 18px 45px rgba(15,23,42,.18);z-index:9999;padding:6px;display:none;}
.autocomplete-item{padding:9px 10px;border-radius:9px;cursor:pointer;font-size:12px;color:#0f172a;display:flex;justify-content:space-between;gap:8px;}
.autocomplete-item:hover,.autocomplete-item.active{background:#eef2ff;color:#3730a3;}
.autocomplete-empty{padding:9px 10px;color:#94a3b8;font-size:12px;}
.check{width:18px;height:18px;accent-color:#4f46e5;} .check:disabled{accent-color:#94a3b8;cursor:not-allowed;}
.pill{display:inline-flex;align-items:center;gap:5px;border-radius:999px;padding:4px 8px;font-size:11px;font-weight:700;white-space:nowrap;}
.pill-ok{background:#dcfce7;color:#166534;}.pill-warn{background:#ffedd5;color:#9a3412;}.pill-skip{background:#e2e8f0;color:#475569;}.pill-dup{background:#e0e7ff;color:#3730a3;}
.skipped-section{display:none;margin-top:14px;border:1px solid #fcd34d;border-radius:10px;overflow:hidden;background:#fff;}
.skipped-header{background:#fffbeb;padding:9px 14px;font-size:12px;font-weight:800;color:#92400e;display:flex;align-items:center;justify-content:space-between;cursor:pointer;user-select:none;}
.skipped-list{max-height:150px;overflow-y:auto;background:#fff;}
.skipped-item{padding:6px 14px;font-size:10px;color:#78350f;border-bottom:1px solid #fef3c7;font-family:monospace;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
</style></head><body>
<h2>📂 Import Bank File — Review Mode</h2>
<div class="subtitle">Choose your bank, or click <b>＋ New Format</b> to register a new CSV layout. Max 7 formats total (3 built-in + 4 custom).</div>

<div class="banks" id="banksContainer">
  <button type="button" class="bank-btn active" id="btn-CIBC" onclick="selectBank('CIBC')"><span class="name">🏦 CIBC</span></button>
  <button type="button" class="bank-btn" id="btn-Amex" onclick="selectBank('Amex')"><span class="name">💳 Amex</span></button>
  <button type="button" class="bank-btn" id="btn-Tangerine" onclick="selectBank('Tangerine')"><span class="name">🍊 Tangerine</span></button>
  <button type="button" class="bank-btn new-fmt" id="btn-new" onclick="openNewFormatInput()"><span class="name">＋ New Format</span></button>
</div>

<!-- Inline name input for new format -->
<div class="name-input-row" id="newFormatNameRow">
  <label>Format name:</label>
  <input type="text" id="newFormatNameInput" placeholder="e.g. LUCAS_AMEXFormat" maxlength="30" onkeydown="if(event.key==='Enter') confirmNewFormatName(); if(event.key==='Escape') cancelNewFormatInput();" />
  <button type="button" class="btn btn-success" id="btnCreateFormat" style="padding:8px 14px;" onclick="confirmNewFormatName()">Create →</button>
  <button type="button" class="btn btn-secondary" id="btnCancelFormat" style="padding:8px 12px;" onclick="cancelNewFormatInput()">Cancel</button>
  <span id="newFormatError" style="font-size:12px;color:#ef4444;"></span>
</div>

<!-- Normal import area -->
<div id="normalArea">
  <div class="dropzone" id="dropzone" onclick="document.getElementById('fileInput').click()">
    <input type="file" id="fileInput" accept=".csv,.xls,.xlsx" onchange="if(this.files[0]) handleFile(this.files[0]);" />
    <div style="font-size:38px;margin-bottom:8px;">📄</div>
    <div id="dz-sub"><b>Click to choose or drag &amp; drop</b><br><span style="font-size:11px;color:#94a3b8">CIBC/Tangerine: .csv · Amex: .xls, .xlsx or .csv</span></div>
  </div>
  <div class="file-selected" id="fileSelected">
    <span>✅</span><span class="fs-name" id="fsName"></span>
    <button type="button" class="fs-remove" id="btnRemoveFile" onclick="removeFile()">✕</button>
  </div>
  <div class="actions">
    <button type="button" class="btn btn-primary" id="btnPreview" disabled onclick="previewImport()">Preview transactions</button>
    <button type="button" class="btn btn-secondary" id="btnReset" onclick="removeFile()">Reset</button>
    <span id="status"></span>
  </div>
</div>

<!-- Setup panel: shown after name is confirmed -->
<div id="setupPanel" class="setup-panel" style="display:none;">
  <div class="setup-hdr">
    <span class="setup-title" id="setupPanelTitle">⚙️ Register New CSV Format</span>
    <button type="button" class="btn btn-secondary" id="btnCancelSetup" style="padding:6px 12px;font-size:12px;" onclick="closeSetupPanel()">✕ Cancel</button>
  </div>
  <div class="setup-row">
    <label class="setup-label">Format Name</label>
    <input type="text" id="setupName" class="setup-input" readonly style="background:#f8fafc;color:#4338ca;font-weight:700;" />
  </div>
  <div class="setup-dropzone" id="setupDropzone">
    <input type="file" id="setupFileInput" accept=".csv" onchange="onSetupFileSelected(event);" />
    <div style="font-size:28px;margin-bottom:6px;">📄</div>
    <div><b>Upload a sample CSV</b> to detect columns</div>
    <div style="font-size:11px;color:#6366f1;margin-top:4px;">Only .csv files supported for setup</div>
  </div>
  <div id="mappingPanel" style="display:none;">
    <div style="font-size:12px;font-weight:700;color:#334155;margin-bottom:8px;">Map each required field to the matching CSV column:</div>
    <div class="mapping-wrap">
      <div class="map-col-left">
        <div class="map-col-hdr">Required Fields</div>
        <div class="map-field-row"><div class="map-field-inner"><span style="font-size:15px;">📅</span><span class="map-field-name">Date <span class="map-field-req">*</span></span></div></div>
        <div class="map-field-row"><div class="map-field-inner"><span style="font-size:15px;">📝</span><span class="map-field-name">Description <span class="map-field-req">*</span></span></div></div>
        <div class="map-field-row"><div class="map-field-inner"><span style="font-size:15px;">💰</span><span class="map-field-name">Amount <span class="map-field-req">*</span></span></div></div>
      </div>
      <div class="map-col-arrow">
        <div class="map-col-hdr" style="justify-content:center;">←</div>
        <div class="map-arrow-row">←</div>
        <div class="map-arrow-row">←</div>
        <div class="map-arrow-row">←</div>
      </div>
      <div class="map-col-right">
        <div class="map-col-hdr">CSV Columns (from your file)</div>
        <div class="map-field-row"><div class="map-right-inner"><select id="mapDate" class="map-select" onchange="renderSamplePreview()"></select><div class="map-sample" id="sampleDate"></div></div></div>
        <div class="map-field-row"><div class="map-right-inner"><select id="mapDesc" class="map-select" onchange="renderSamplePreview()"></select><div class="map-sample" id="sampleDesc"></div></div></div>
        <div class="map-field-row"><div class="map-right-inner"><select id="mapAmount" class="map-select" onchange="renderSamplePreview()"></select><div class="map-sample" id="sampleAmount"></div></div></div>
      </div>
    </div>
    <div class="setup-opts">
      <div class="setup-opt">
        <label>Date Format</label>
        <select id="setupDateFormat">
          <option value="auto">Auto-detect</option>
          <option value="dd MMM yyyy">dd MMM yyyy &nbsp;(09 Jun 2026)</option>
          <option value="yyyy-MM-dd">yyyy-MM-dd &nbsp;(2026-06-09)</option>
          <option value="MM/dd/yyyy">MM/dd/yyyy &nbsp;(06/09/2026)</option>
          <option value="dd/MM/yyyy">dd/MM/yyyy &nbsp;(09/06/2026)</option>
        </select>
      </div>
      <div class="setup-opt">
        <label>Expense Amount Sign</label>
        <select id="setupAmountSign">
          <option value="positive">Positive = Expense &nbsp;(25.00)</option>
          <option value="negative">Negative = Expense &nbsp;(-25.00)</option>
        </select>
      </div>
    </div>
    <div class="sample-wrap">
      <div class="sample-lbl">📋 Preview — first rows with your mapping:</div>
      <div id="samplePreview"></div>
    </div>
    <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;">
      <button type="button" class="btn btn-success" id="btnSaveFormat" onclick="saveFormatAndPreview()">💾 Save &amp; Preview Transactions</button>
      <button type="button" class="btn btn-secondary" id="btnCancelSetup2" onclick="closeSetupPanel()">Cancel</button>
      <span id="setupStatus" style="font-size:12px;flex:1;"></span>
    </div>
  </div>
</div>

<!-- Review panel -->
<div class="review-panel" id="reviewPanel">
  <div class="review-header">
    <div>
      <div class="review-title" id="reviewTitle">Review transactions</div>
      <div class="tiny" id="reviewSummary"></div>
    </div>
    <div class="review-tools">
      <input class="search" id="filterBox" placeholder="Search description/category..." oninput="renderReviewTable()" />
      <button type="button" class="btn btn-secondary" id="btnCheckAll" onclick="setAllInclude(true)">Check all</button>
      <button type="button" class="btn btn-secondary" id="btnUncheckAll" onclick="setAllInclude(false)">Uncheck all</button>
    </div>
  </div>
  <div class="table-wrap">
    <table>
      <thead><tr><th>Import</th><th>Date</th><th>Description</th><th>Amount</th><th>Bank</th><th>Category</th><th>Status</th></tr></thead>
      <tbody id="reviewBody"></tbody>
    </table>
  </div>
  <div class="actions" style="padding:12px 14px;border-top:1px solid #e2e8f0;margin-top:0;">
    <button type="button" class="btn btn-primary" id="btnCommit" onclick="commitImport()">Confirm import selected</button>
    <button type="button" class="btn btn-danger" id="btnCancelImport" onclick="removeFile()">Cancel</button>
    <span class="tiny" id="commitHint"></span>
  </div>
</div>
<div class="skipped-section" id="skipped-section">
  <div class="skipped-header" id="skippedHeader" onclick="toggleSkipped()">
    <span id="skipped-title">⏭ Skipped duplicates</span>
    <span id="skipped-toggle">▼ Show</span>
  </div>
  <div id="skipped-list" class="skipped-list" style="display:none"></div>
</div>

<script>
  // ── State ──
  var selectedBank = 'CIBC';
  var selectedFile = null;
  var selectedCustomFormat = null;
  var amexParsedRows = null;
  var reviewRows = [];
  var categoryOptions = [];
  var customFormats = [];
  var setupCsvHeaders = [];
  var setupCsvSampleRows = [];
  var setupFile = null;
  var MAX_CUSTOM = 4;

  // ── Init: load saved formats + drag-drop ──
  google.script.run
    .withSuccessHandler(function(res) {
      if (res && res.ok && res.formats && res.formats.length > 0) {
        customFormats = res.formats;
        renderBankButtons();
      }
    })
    .withFailureHandler(function(){})
    .getCsvFormats();

  // Drag & Drop (cannot be inline)
  (function() {
    var dz = document.getElementById('dropzone');
    if (!dz) return;
    ['dragenter','dragover'].forEach(function(ev){ dz.addEventListener(ev, function(e){ e.preventDefault(); dz.classList.add('drag'); }); });
    ['dragleave','drop'].forEach(function(ev){ dz.addEventListener(ev, function(e){ e.preventDefault(); dz.classList.remove('drag'); }); });
    dz.addEventListener('drop', function(e){ var f = e.dataTransfer.files && e.dataTransfer.files[0]; if(f) handleFile(f); });
  })();

  // ── Helpers ──
  function esc(v){ return String(v==null?'':v).replace(/[&<>"']/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }
  function money(n){ return '$'+Number(n||0).toLocaleString('en-CA',{minimumFractionDigits:2,maximumFractionDigits:2}); }

  // ── Bank buttons ──
  function renderBankButtons() {
    document.querySelectorAll('.bank-btn-custom').forEach(function(b){ b.remove(); });
    var container = document.getElementById('banksContainer');
    var newBtn = document.getElementById('btn-new');
    customFormats.forEach(function(fmt) {
      var btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'bank-btn custom-fmt bank-btn-custom';
      btn.id = 'btn-cf-' + fmt.name;
      var nameSpan = document.createElement('span');
      nameSpan.className = 'name';
      nameSpan.textContent = '📋 ' + fmt.name;
      btn.appendChild(nameSpan);
      var delSpan = document.createElement('span');
      delSpan.className = 'del-fmt';
      delSpan.title = 'Delete this format';
      delSpan.textContent = '×';
      delSpan.addEventListener('click', (function(n){ return function(e){ e.stopPropagation(); deleteFormat(n); }; })(fmt.name));
      btn.appendChild(delSpan);
      btn.addEventListener('click', (function(f){ return function(){ selectCustomFormat(f); }; })(fmt));
      container.insertBefore(btn, newBtn);
    });
    newBtn.style.display = customFormats.length >= MAX_CUSTOM ? 'none' : '';
  }

  function selectBank(b) {
    selectedBank = b; selectedCustomFormat = null;
    document.querySelectorAll('.bank-btn').forEach(function(x){ x.classList.remove('active'); });
    var btn = document.getElementById('btn-' + b); if (btn) btn.classList.add('active');
    closeSetupPanel(); cancelNewFormatInput(); removeFile(false);
    document.getElementById('dz-sub').innerHTML = '<b>Click to choose or drag &amp; drop</b><br><span style="font-size:11px;color:#94a3b8">CIBC/Tangerine: .csv · Amex: .xls, .xlsx or .csv</span>';
  }

  function selectCustomFormat(fmt) {
    selectedBank = fmt.name; selectedCustomFormat = fmt;
    document.querySelectorAll('.bank-btn').forEach(function(x){ x.classList.remove('active'); });
    var btn = document.getElementById('btn-cf-' + fmt.name); if (btn) btn.classList.add('active');
    closeSetupPanel(); cancelNewFormatInput(); removeFile(false);
    document.getElementById('dz-sub').innerHTML = '<b>Click to choose or drag &amp; drop</b><br><span style="font-size:11px;color:#6366f1">Format: <b>' + esc(fmt.name) + '</b> · .csv only</span>';
  }

  // ── New Format name-first flow ──
  function openNewFormatInput() {
    document.getElementById('newFormatError').textContent = '';
    document.getElementById('newFormatNameInput').value = '';
    document.getElementById('newFormatNameRow').style.display = 'flex';
    document.getElementById('btn-new').style.display = 'none';
    document.getElementById('newFormatNameInput').focus();
  }

  function cancelNewFormatInput() {
    document.getElementById('newFormatNameRow').style.display = 'none';
    if (customFormats.length < MAX_CUSTOM) document.getElementById('btn-new').style.display = '';
  }

  function confirmNewFormatName() {
    var name = document.getElementById('newFormatNameInput').value.trim();
    var errEl = document.getElementById('newFormatError');
    if (!name) { errEl.textContent = 'Please enter a name.'; return; }
    if (['CIBC','Amex','Tangerine'].indexOf(name) >= 0) { errEl.textContent = '"' + name + '" is reserved.'; return; }
    if (customFormats.some(function(f){ return f.name === name; })) { errEl.textContent = '"' + name + '" already exists.'; return; }
    cancelNewFormatInput();
    openSetupPanel(name);
  }

  // ── Delete format ──
  function deleteFormat(name) {
    if (!confirm('Delete format "' + name + '"?\nThis cannot be undone.')) return;
    google.script.run
      .withSuccessHandler(function(res) {
        if (!res || !res.ok) { alert('Error: ' + (res ? res.error : 'unknown')); return; }
        customFormats = customFormats.filter(function(f){ return f.name !== name; });
        if (selectedCustomFormat && selectedCustomFormat.name === name) {
          selectedCustomFormat = null; selectedBank = 'CIBC'; removeFile(false);
        }
        renderBankButtons();
        document.querySelectorAll('.bank-btn').forEach(function(x){ x.classList.remove('active'); });
        var active = document.getElementById('btn-' + selectedBank) || document.getElementById('btn-cf-' + selectedBank);
        if (active) active.classList.add('active');
      })
      .withFailureHandler(function(e){ alert('Error: ' + (e.message||e)); })
      .deleteCsvFormat(name);
  }

  // ── Normal file handling ──
  function handleFile(file) {
    selectedFile = file; amexParsedRows = null; reviewRows = [];
    document.getElementById('fsName').textContent = file.name;
    document.getElementById('fileSelected').style.display = 'flex';
    document.getElementById('dropzone').style.display = 'none';
    document.getElementById('btnPreview').disabled = false;
    document.getElementById('reviewPanel').style.display = 'none';
    document.getElementById('skipped-section').style.display = 'none';
    document.getElementById('status').className = 'inf';
    document.getElementById('status').textContent = 'Ready to preview.';
    if (selectedBank === 'Amex' && (file.name.toLowerCase().endsWith('.xls') || file.name.toLowerCase().endsWith('.xlsx'))) {
      parseAmexXlsInBrowser(file);
    }
  }

  function removeFile(clearInput) {
    if (clearInput !== false) { var fi = document.getElementById('fileInput'); if(fi) fi.value = ''; }
    selectedFile = null; amexParsedRows = null; reviewRows = []; categoryOptions = [];
    document.getElementById('fileSelected').style.display = 'none';
    document.getElementById('dropzone').style.display = 'block';
    document.getElementById('btnPreview').disabled = true;
    document.getElementById('status').textContent = '';
    document.getElementById('reviewPanel').style.display = 'none';
    document.getElementById('skipped-section').style.display = 'none';
  }

  // ── Setup Panel ──
  function openSetupPanel(name) {
    document.querySelectorAll('.bank-btn').forEach(function(x){ x.classList.remove('active'); });
    document.getElementById('normalArea').style.display = 'none';
    document.getElementById('setupPanel').style.display = 'block';
    document.getElementById('reviewPanel').style.display = 'none';
    document.getElementById('mappingPanel').style.display = 'none';
    document.getElementById('setupName').value = name || '';
    document.getElementById('setupPanelTitle').textContent = '⚙️ Register: ' + (name || 'New Format');
    document.getElementById('setupFileInput').value = '';
    document.getElementById('setupStatus').textContent = '';
    setupFile = null; setupCsvHeaders = []; setupCsvSampleRows = [];
  }

  function closeSetupPanel() {
    document.getElementById('setupPanel').style.display = 'none';
    document.getElementById('normalArea').style.display = 'block';
    document.querySelectorAll('.bank-btn').forEach(function(x){ x.classList.remove('active'); });
    var btn = document.getElementById('btn-' + selectedBank) || document.getElementById('btn-cf-' + selectedBank);
    if (btn) btn.classList.add('active');
  }

  function onSetupFileSelected(event) {
    var file = event.target.files[0]; if (!file) return;
    setupFile = file;
    var reader = new FileReader();
    reader.onload = function(e){ parseSetupCsvBrowser(e.target.result); };
    reader.readAsText(file);
  }

  function parseCsvLineBrowser(line) {
    var result = [], cur = '', inQ = false;
    for (var i = 0; i < line.length; i++) {
      var ch = line[i], nx = line[i+1];
      if (ch === '"') { if (inQ && nx === '"') { cur += '"'; i++; } else inQ = !inQ; }
      else if (ch === ',' && !inQ) { result.push(cur.trim()); cur = ''; }
      else cur += ch;
    }
    result.push(cur.trim()); return result;
  }

  function parseSetupCsvBrowser(text) {
    var lines = text.replace(/\r\n/g,'\n').replace(/\r/g,'\n').split('\n').filter(function(l){ return l.trim(); });
    if (lines.length < 2) { document.getElementById('setupStatus').className='err'; document.getElementById('setupStatus').textContent='❌ CSV must have header + at least one data row.'; return; }
    setupCsvHeaders = parseCsvLineBrowser(lines[0]);
    setupCsvSampleRows = [];
    for (var i = 1; i < Math.min(4, lines.length); i++) setupCsvSampleRows.push(parseCsvLineBrowser(lines[i]));
    renderMappingPanel();
  }

  function renderMappingPanel() {
    ['mapDate','mapDesc','mapAmount'].forEach(function(id) {
      var sel = document.getElementById(id); sel.innerHTML = '';
      setupCsvHeaders.forEach(function(h, idx) { var o = document.createElement('option'); o.value=idx; o.textContent=h; sel.appendChild(o); });
    });
    setupCsvHeaders.forEach(function(h, idx) {
      var l = h.toLowerCase().replace(/[^a-z]/g,'');
      if (l==='date'||(l.includes('date')&&!l.includes('process')&&!l.includes('post'))) document.getElementById('mapDate').value=idx;
      if (['description','desc','narrative','memo','name','details','merchant'].some(function(k){ return l.includes(k); })) document.getElementById('mapDesc').value=idx;
      if (['amount','amt','debit','value','total'].some(function(k){ return l.includes(k); })) document.getElementById('mapAmount').value=idx;
    });
    renderSamplePreview();
    document.getElementById('mappingPanel').style.display = 'block';
  }

  function renderSamplePreview() {
    var di=parseInt(document.getElementById('mapDate').value), ni=parseInt(document.getElementById('mapDesc').value), ai=parseInt(document.getElementById('mapAmount').value);
    var s = setupCsvSampleRows[0]||[];
    document.getElementById('sampleDate').textContent   = s[di] ? 'e.g. '+s[di]  : '';
    document.getElementById('sampleDesc').textContent   = s[ni] ? 'e.g. '+s[ni]  : '';
    document.getElementById('sampleAmount').textContent = s[ai] ? 'e.g. '+s[ai]  : '';
    var dh=esc(setupCsvHeaders[di]||'?'), nh=esc(setupCsvHeaders[ni]||'?'), ah=esc(setupCsvHeaders[ai]||'?');
    var html='<table><thead><tr><th>Date ('+dh+')</th><th>Description ('+nh+')</th><th>Amount ('+ah+')</th></tr></thead><tbody>';
    setupCsvSampleRows.forEach(function(row){ html+='<tr><td>'+esc(row[di]||'')+'</td><td>'+esc(row[ni]||'')+'</td><td>'+esc(row[ai]||'')+'</td></tr>'; });
    html+='</tbody></table>';
    document.getElementById('samplePreview').innerHTML = html;
  }

  function saveFormatAndPreview() {
    var name=document.getElementById('setupName').value.trim(), statusEl=document.getElementById('setupStatus');
    if (!name) { statusEl.className='err'; statusEl.textContent='❌ Format name is required.'; return; }
    if (!setupFile) { statusEl.className='err'; statusEl.textContent='❌ Please upload a CSV file.'; return; }
    var config = { name:name, dateCol:parseInt(document.getElementById('mapDate').value), descCol:parseInt(document.getElementById('mapDesc').value),
      amountCol:parseInt(document.getElementById('mapAmount').value), dateFormat:document.getElementById('setupDateFormat').value,
      amountSign:document.getElementById('setupAmountSign').value, skipRows:1 };
    statusEl.className='inf'; statusEl.textContent='⏳ Saving format...';
    google.script.run
      .withSuccessHandler(function(res) {
        if (!res.ok) { statusEl.className='err'; statusEl.textContent='❌ '+res.error; return; }
        var existing=customFormats.findIndex(function(f){ return f.name===config.name; });
        if (existing>=0) customFormats[existing]=config; else customFormats.push(config);
        renderBankButtons();
        selectedCustomFormat=config; selectedBank=config.name; selectedFile=setupFile;
        closeSetupPanel();
        document.getElementById('fsName').textContent=setupFile.name;
        document.getElementById('fileSelected').style.display='flex';
        document.getElementById('dropzone').style.display='none';
        document.getElementById('btnPreview').disabled=false;
        document.querySelectorAll('.bank-btn').forEach(function(x){ x.classList.remove('active'); });
        var cfBtn=document.getElementById('btn-cf-'+config.name); if(cfBtn) cfBtn.classList.add('active');
        document.getElementById('status').className='ok'; document.getElementById('status').textContent='✅ Format "'+esc(name)+'" saved!';
        previewImport();
      })
      .withFailureHandler(function(e){ statusEl.className='err'; statusEl.textContent='❌ '+(e.message||e); })
      .saveCsvFormat(config);
  }

  // ── Amex XLS ──
  function parseAmexXlsInBrowser(file) {
    document.getElementById('status').className='inf'; document.getElementById('status').textContent='⏳ Reading Amex file...';
    var reader = new FileReader();
    reader.onload = function(e) {
      try {
        var wb=XLSX.read(new Uint8Array(e.target.result),{type:'array',cellDates:true}), ws=wb.Sheets[wb.SheetNames[0]];
        var rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''}), headerIdx=-1;
        for (var i=0;i<rows.length;i++){ if(String(rows[i][0]||'').trim()==='Date'){headerIdx=i;break;} }
        if (headerIdx===-1) throw new Error('Could not find Amex transaction table.');
        var parsed=[];
        for (var i=headerIdx+1;i<rows.length;i++) {
          var row=rows[i],rawDate=row[0],desc=String(row[2]||'').trim(),rawAmt=row[3];
          if (!rawDate||!desc) continue;
          var dateStr;
          if (rawDate instanceof Date){var M=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];dateStr=rawDate.getDate()+' '+M[rawDate.getMonth()]+'. '+rawDate.getFullYear();}
          else dateStr=String(rawDate).trim();
          var amtStr=String(rawAmt||'').replace(/[$,]/g,'').trim();
          if (!dateStr||isNaN(parseFloat(amtStr))) continue;
          parsed.push({date:dateStr,desc:desc,amount:amtStr});
        }
        if (parsed.length===0) throw new Error('No Amex transactions found.');
        amexParsedRows=parsed;
        document.getElementById('status').className='ok'; document.getElementById('status').textContent='✅ Amex file read. Click Preview transactions.';
      } catch(err){ document.getElementById('status').className='err'; document.getElementById('status').textContent='❌ '+err.message; document.getElementById('btnPreview').disabled=true; }
    };
    reader.readAsArrayBuffer(file);
  }

  // ── Preview ──
  function previewImport() {
    if (!selectedFile) return;
    var status=document.getElementById('status');
    status.className='inf'; status.textContent='⏳ Building preview...'; document.getElementById('btnPreview').disabled=true;
    function onPreview(res) {
      document.getElementById('btnPreview').disabled=false;
      if (!res.ok){status.className='err';status.textContent='❌ '+res.error;return;}
      reviewRows=res.transactions||[]; categoryOptions=res.categories||[];
      status.className='ok'; status.textContent='✅ Preview ready. Duplicates are unchecked automatically.';
      document.getElementById('reviewPanel').style.display='block'; renderReviewTable();
    }
    function fail(e){status.className='err';status.textContent='❌ '+(e.message||e);document.getElementById('btnPreview').disabled=false;}
    var reader=new FileReader();
    if (selectedCustomFormat) {
      reader.onload=function(e){ google.script.run.withSuccessHandler(onPreview).withFailureHandler(fail).previewFileImport(selectedBank,e.target.result.split(',')[1],selectedFile.name,selectedCustomFormat); };
      reader.readAsDataURL(selectedFile); return;
    }
    if (selectedBank==='Amex'&&amexParsedRows) {
      google.script.run.withSuccessHandler(onPreview).withFailureHandler(fail).previewFileImport('Amex',JSON.stringify({rows:amexParsedRows}),selectedFile.name); return;
    }
    reader.onload=function(e){ google.script.run.withSuccessHandler(onPreview).withFailureHandler(fail).previewFileImport(selectedBank,e.target.result.split(',')[1],selectedFile.name); };
    reader.readAsDataURL(selectedFile);
  }

  // ── Review Table ──
  function getReviewSortRank(row){ if(row&&row.include&&!row.isDuplicate)return 0; if(row&&!row.include&&!row.isDuplicate)return 1; return 2; }
  function getSortedReviewEntries(){
    return reviewRows.map(function(row,index){ return {row:row,index:index}; }).sort(function(a,b){
      var ra=getReviewSortRank(a.row),rb=getReviewSortRank(b.row); if(ra!==rb)return ra-rb;
      var da=String(a.row.dateISO||''),db=String(b.row.dateISO||''); if(da!==db)return db.localeCompare(da);
      return Number(b.row.amount||0)-Number(a.row.amount||0);
    });
  }

  function renderReviewTable() {
    var q=String(document.getElementById('filterBox').value||'').toLowerCase(), body=document.getElementById('reviewBody');
    var selected=0,total=0,uncat=0,dups=0,sum=0, html=[];
    getSortedReviewEntries().forEach(function(entry) {
      var r=entry.row,i=entry.index;
      var hay=(r.desc+' '+r.category+' '+r.bank+' '+r.dateLabel+' '+(r.isDuplicate?'duplicate':'')).toLowerCase();
      if(q&&hay.indexOf(q)<0)return; total++;
      if(r.isDuplicate)dups++; if(r.include){selected++;sum+=Number(r.amount||0);} if(r.include&&!r.category&&!r.isDuplicate)uncat++;
      var cls=(r.include?'':'excluded ')+(r.isDuplicate?'duplicate ':'')+((!r.category&&!r.isDuplicate)?'uncat':'');
      var st=r.isDuplicate?'<span class="pill pill-dup">🔁 Duplicate</span>':(!r.include?'<span class="pill pill-skip">Skipped</span>':(!r.category?'<span class="pill pill-warn">Needs category</span>':'<span class="pill pill-ok">Ready</span>'));
      html.push('<tr class="'+cls+'"><td><input class="check" type="checkbox" '+(r.include?'checked':'')+' '+(r.isDuplicate?'disabled title="Already imported"':'')+' onchange="updateInclude('+i+',this.checked)"></td><td>'+esc(r.dateLabel)+'</td><td class="desc" title="'+esc(r.desc)+'">'+esc(r.desc)+'</td><td class="amount">'+money(r.amount)+'</td><td>'+esc(r.bank)+'</td><td class="cat-cell"><input autocomplete="off" class="cat-input '+(!r.category&&!r.isDuplicate?'missing':'')+'" id="cat-'+i+'" value="'+esc(r.category)+'" placeholder="Type category..." oninput="updateCategorySoft('+i+',this.value)" onfocus="openCategoryAutocomplete('+i+')" onkeydown="categoryKeydown(event,'+i+')"><div class="autocomplete-menu" id="cat-menu-'+i+'"></div></td><td>'+st+'</td></tr>');
    });
    body.innerHTML=html.join('')||'<tr><td colspan="7" class="tiny">No matching rows.</td></tr>';
    document.getElementById('reviewTitle').textContent='Review '+reviewRows.length+' parsed transactions';
    document.getElementById('reviewSummary').textContent=selected+' selected · '+dups+' duplicate unchecked · '+uncat+' selected without category · '+money(sum)+' selected total';
    document.getElementById('commitHint').textContent='Duplicates are unchecked by default. Only checked non-duplicate rows will be imported.';
    document.getElementById('btnCommit').disabled=selected===0;
  }

  function updateInclude(i,v){ if(!reviewRows[i])return; if(reviewRows[i].isDuplicate){reviewRows[i].include=false;renderReviewTable();return;} reviewRows[i].include=v; renderReviewTable(); }
  function updateCategorySoft(i,v){ if(!reviewRows[i])return; reviewRows[i].category=String(v||'').trim(); renderCategorySuggestions(i); updateReviewSummaryOnly(); }
  function updateReviewSummaryOnly(){
    var selected=0,uncat=0,dups=0,sum=0;
    reviewRows.forEach(function(r){ if(r.isDuplicate)dups++; if(r.include){selected++;sum+=Number(r.amount||0);} if(r.include&&!r.category&&!r.isDuplicate)uncat++; });
    document.getElementById('reviewSummary').textContent=selected+' selected · '+dups+' duplicate unchecked · '+uncat+' selected without category · '+money(sum)+' selected total';
    document.getElementById('btnCommit').disabled=selected===0;
  }
  function setAllInclude(v){ reviewRows.forEach(function(r){ r.include=v?!r.isDuplicate:false; }); renderReviewTable(); }
  function openCategoryAutocomplete(i){ renderCategorySuggestions(i); }
  function closeAllAutocomplete(exceptId){ document.querySelectorAll('.autocomplete-menu').forEach(function(m){ if(!exceptId||m.id!==exceptId)m.style.display='none'; }); }
  function renderCategorySuggestions(i){
    var input=document.getElementById('cat-'+i),menu=document.getElementById('cat-menu-'+i);
    if(!input||!menu)return; closeAllAutocomplete(menu.id);
    var val=String(input.value||'').trim().toLowerCase(), list=categoryOptions.slice();
    if(val){ list=list.filter(function(c){return c.toLowerCase().indexOf(val)!==-1;}).sort(function(a,b){ var av=a.toLowerCase(),bv=b.toLowerCase(); var ap=av.startsWith(val)?0:1,bp=bv.startsWith(val)?0:1; if(ap!==bp)return ap-bp; return av.localeCompare(bv); }); }
    list=list.slice(0,10);
    if(list.length===0){menu.innerHTML='<div class="autocomplete-empty">No category found. You can type a new one.</div>';menu.style.display='block';return;}
    menu.innerHTML='';
    list.forEach(function(c,idx){
      var item=document.createElement('div'); item.className='autocomplete-item'+(idx===0?' active':''); item.setAttribute('data-value',c);
      var label=document.createElement('span'); label.textContent=c; item.appendChild(label);
      item.addEventListener('mousedown',function(ev){ev.preventDefault();selectCategorySuggestion(i,c);});
      menu.appendChild(item);
    }); menu.style.display='block';
  }
  function selectCategorySuggestion(i,value){
    if(!reviewRows[i])return; var v=String(value||'').trim(); reviewRows[i].category=v;
    var input=document.getElementById('cat-'+i); if(input)input.value=v;
    closeAllAutocomplete(); renderReviewTable();
  }
  function categoryKeydown(e,i){
    var menu=document.getElementById('cat-menu-'+i); if(!menu||menu.style.display==='none')return;
    var items=Array.prototype.slice.call(menu.querySelectorAll('.autocomplete-item')); if(!items.length)return;
    var idx=items.findIndex(function(x){ return x.classList.contains('active'); }); if(idx<0)idx=0;
    if(e.key==='ArrowDown'){e.preventDefault();if(items[idx])items[idx].classList.remove('active');idx=(idx+1)%items.length;items[idx].classList.add('active');items[idx].scrollIntoView({block:'nearest'});}
    else if(e.key==='ArrowUp'){e.preventDefault();if(items[idx])items[idx].classList.remove('active');idx=(idx-1+items.length)%items.length;items[idx].classList.add('active');items[idx].scrollIntoView({block:'nearest'});}
    else if(e.key==='Enter'){e.preventDefault();selectCategorySuggestion(i,items[Math.max(0,idx)].getAttribute('data-value'));}
    else if(e.key==='Escape'){closeAllAutocomplete();}
  }
  document.addEventListener('click',function(e){ if(!e.target.classList||!e.target.classList.contains('cat-input'))closeAllAutocomplete(); });

  // ── Commit ──
  function commitImport(){
    var status=document.getElementById('status');
    status.className='inf'; status.textContent='⏳ Importing selected transactions...'; document.getElementById('btnCommit').disabled=true;
    google.script.run.withSuccessHandler(function(res){
      document.getElementById('btnCommit').disabled=false;
      if(res.ok){
        var addedTxt=res.added>0?'✅ Added: '+res.added:'⚠️ Nothing new';
        var skipTxt=res.skipped>0?'  ·  ⏭ Skipped: '+res.skipped+' duplicates':'';
        status.className=res.added>0?'ok':'warn'; status.textContent=addedTxt+skipTxt;
        if(res.skipped>0&&res.skippedList&&res.skippedList.length>0){
          document.getElementById('skipped-title').textContent='⏭ Skipped '+res.skipped+' duplicates';
          document.getElementById('skipped-list').innerHTML=res.skippedList.map(function(s){return '<div class="skipped-item">'+esc(s)+'</div>';}).join('');
          document.getElementById('skipped-section').style.display='block';
        }
        if(res.added>0)document.getElementById('reviewPanel').style.display='none';
      } else {status.className='err';status.textContent='❌ '+res.error;}
    }).withFailureHandler(function(e){status.className='err';status.textContent='❌ '+(e.message||e);document.getElementById('btnCommit').disabled=false;})
    .commitReviewedImport(selectedBank, reviewRows);
  }

  var skippedOpen=false;
  function toggleSkipped(){ skippedOpen=!skippedOpen; document.getElementById('skipped-list').style.display=skippedOpen?'block':'none'; document.getElementById('skipped-toggle').textContent=skippedOpen?'▲ Hide':'▼ Show'; }
<\/script></body></html>`;
}

// Compatibility wrapper.
// The menu uses buildReviewImportModalHTML_(), but this prevents old calls from breaking.
function buildModalHTML() {
  return buildReviewImportModalHTML_();
}
