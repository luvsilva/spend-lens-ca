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
    const rs = ss.getSheetByName(CONFIG.RULES_SHEET);
    if (rs && rs.getLastRow() > 1) {
      const skipPfx = ['─','—','➕','ℹ','#','=','—','✍'];
      rs.getRange(2, 2, rs.getLastRow() - 1, 1).getValues().forEach(function(r) {
        const c = String(r[0] || "").trim();
        if (c && !CONFIG.INCOME_CATEGORIES.includes(c) && !skipPfx.some(function(p){ return c.startsWith(p); })) set.add(c);
      });
    }
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
  ensureCategoryRulesSeeded_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rs = ss.getSheetByName(CONFIG.RULES_SHEET);
  if (!rs || rs.getLastRow() < 2) return [];
  const data = rs.getRange(2, 1, rs.getLastRow()-1, 2).getValues();
  const rules = [];
  const skipPrefixes = ['─','—','➕','ℹ','#','=','—','✍'];
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

  // Preserve custom/learned rules before clearing
  const preservedCustom = [];
  if (rs && rs.getLastRow() >= 2) {
    const data = rs.getRange(2, 1, rs.getLastRow() - 1, 2).getValues();
    const skipPrefixes_ = ['─','—','➕','ℹ','#','=','✍'];
    let inCustom = false;
    data.forEach(function(row) {
      const kw  = String(row[0] || '').trim();
      const cat = String(row[1] || '').trim();
      if (kw.startsWith('✍')) { inCustom = true; return; }
      if (!inCustom || !kw || !cat) return;
      if (skipPrefixes_.some(function(p) { return kw.startsWith(p); })) return;
      preservedCustom.push({ keyword: kw, category: cat });
    });
  }

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

  const ROW_COLORS = ["#ffffff","#f8fafc"];
  const SECTION_BG = "#e2e8f0";
  const SECTION_FG = "#334155";

  let row = 3;
  order.forEach(function(cat, gi) {
    const keywords = grouped[cat];
    if (!keywords || keywords.length === 0) return;
    const emoji    = EMOJIS[cat] || "📌";
    const rowColor = ROW_COLORS[gi % 2];

    rs.getRange(row, 1, 1, 2).merge()
      .setValue(emoji + "   " + cat.toUpperCase())
      .setBackground(SECTION_BG).setFontColor(SECTION_FG)
      .setFontWeight("bold").setFontSize(10).setVerticalAlignment("middle");
    rs.setRowHeight(row, 26);
    row++;

    const vals = keywords.map(function(kw) { return [kw, cat]; });
    rs.getRange(row, 1, vals.length, 2).setValues(vals)
      .setBackground(rowColor).setFontSize(11);
    row += vals.length;
  });

  // Learn new rules and merge with preserved custom rules
  const learnedRules = learnNewCategoryRules_(preservedCustom);
  const allCustom = mergeCustomRules_(preservedCustom, learnedRules);

  // CUSTOM & LEARNED section header (always present so the user knows where to add rules)
  rs.getRange(row, 1, 1, 2).merge()
    .setValue("✍  CUSTOM & LEARNED RULES")
    .setBackground("#fef3c7").setFontColor("#92400e")
    .setFontWeight("bold").setFontSize(10).setVerticalAlignment("middle");
  rs.setRowHeight(row, 26);
  row++;

  if (allCustom.length > 0) {
    const vals = allCustom.map(function(r) { return [r.keyword, r.category]; });
    rs.getRange(row, 1, vals.length, 2).setValues(vals)
      .setBackground("#fffbeb").setFontSize(11);
  }

  rs.setColumnWidth(1, 300);
  rs.setColumnWidth(2, 190);
  if (rs.isSheetHidden()) rs.showSheet();
  ss.setActiveSheet(rs);

  if (learnedRules.length > 0) {
    ss.toast(learnedRules.length + ' new rule(s) learned from your transaction history.', '🧠 Rules Updated', 5);
  }
}

function toggleCategoryRulesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rs = ss.getSheetByName(CONFIG.RULES_SHEET);
  if (rs && !rs.isSheetHidden()) { rs.hideSheet(); return; }
  setupCategoryRulesSheet();
}

// Scans Expenses_Form for transactions with a category that no existing rule covers.
// Returns an array of { keyword, category } for consistent new rules only.
function learnNewCategoryRules_(existingCustomRules) {
  const knownKeywords = new Set();
  (CONFIG.CATEGORIES || []).forEach(function(rule) {
    (rule.match || []).forEach(function(kw) { knownKeywords.add(kw.toUpperCase()); });
  });
  (existingCustomRules || []).forEach(function(r) { knownKeywords.add(r.keyword.toUpperCase()); });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expSheet = ss.getSheetByName(CONFIG.TARGET_SHEET);
  if (!expSheet || expSheet.getLastRow() < 2) return [];

  const data = expSheet.getRange(2, CONFIG.COLS.DESCRIPTION, expSheet.getLastRow() - 1, 2).getValues();
  const tokenCatMap = {};

  data.forEach(function(row) {
    const desc = String(row[0] || '').trim();
    const cat  = String(row[1] || '').trim();
    if (!desc || !cat) return;
    if (CONFIG.INCOME_CATEGORIES.includes(cat)) return;

    const upper = desc.toUpperCase();
    const alreadyMatched =
      (CONFIG.CATEGORIES || []).some(function(rule) {
        return (rule.match || []).some(function(kw) { return upper.includes(kw.toUpperCase()); });
      }) ||
      (existingCustomRules || []).some(function(r) { return upper.includes(r.keyword.toUpperCase()); });
    if (alreadyMatched) return;

    const tokens = extractMeaningfulTokens(desc);
    const single = tokens.filter(function(t) { return !t.includes(' ') && t.length >= 4; });
    const candidates = single.length > 0 ? single : tokens.filter(function(t) { return !t.includes(' '); });
    if (candidates.length === 0) return;
    const token = candidates[0];

    if (!tokenCatMap[token]) tokenCatMap[token] = {};
    tokenCatMap[token][cat] = (tokenCatMap[token][cat] || 0) + 1;
  });

  const learned = [];
  Object.keys(tokenCatMap).forEach(function(token) {
    const cats = Object.keys(tokenCatMap[token]);
    if (cats.length !== 1) return; // ambiguous — skip
    if (knownKeywords.has(token.toUpperCase())) return;
    learned.push({ keyword: token, category: cats[0] });
  });
  return learned;
}

// Merges preserved + learned arrays, deduplicating by keyword (case-insensitive).
function mergeCustomRules_(existing, learned) {
  const seen = new Set((existing || []).map(function(r) { return r.keyword.toUpperCase(); }));
  const merged = (existing || []).slice();
  (learned || []).forEach(function(r) {
    if (!seen.has(r.keyword.toUpperCase())) {
      merged.push(r);
      seen.add(r.keyword.toUpperCase());
    }
  });
  return merged;
}

// Seeds the CategoryRules sheet once from CONFIG.CATEGORIES if the sheet does not yet exist.
// Idempotent: does nothing if the sheet already exists.
function ensureCategoryRulesSeeded_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(CONFIG.RULES_SHEET)) return;

  const rs = ss.insertSheet(CONFIG.RULES_SHEET);

  rs.getRange(1,1,1,2)
    .setValues([["Keyword  —  matches text in your transaction description","Category  —  label applied in the spreadsheet"]])
    .setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold").setFontSize(11);
  rs.setRowHeight(1, 32);

  rs.getRange(2,1,1,2).merge()
    .setValue("ℹ️   To hide this sheet: Budget Importer menu  →  📋 Category Rules")
    .setBackground("#e0f2fe").setFontColor("#0369a1").setFontStyle("italic")
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  rs.setRowHeight(2, 26);
  rs.setFrozenRows(2);

  const grouped = {}, order = [];
  (CONFIG.CATEGORIES || []).forEach(function(rule) {
    if (!grouped[rule.category]) { grouped[rule.category] = []; order.push(rule.category); }
    (rule.match || []).forEach(function(kw) { grouped[rule.category].push(kw); });
  });

  const EMOJIS_S = {
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

  const ROW_COLORS_S = ["#ffffff","#f8fafc"];
  const SECTION_BG_S = "#e2e8f0", SECTION_FG_S = "#334155";
  let row = 3;

  order.forEach(function(cat, gi) {
    const keywords = grouped[cat];
    if (!keywords || keywords.length === 0) return;
    const emoji = EMOJIS_S[cat] || "📌";
    const rowColor = ROW_COLORS_S[gi % 2];

    rs.getRange(row, 1, 1, 2).merge()
      .setValue(emoji + "   " + cat.toUpperCase())
      .setBackground(SECTION_BG_S).setFontColor(SECTION_FG_S)
      .setFontWeight("bold").setFontSize(10).setVerticalAlignment("middle");
    rs.setRowHeight(row, 26);
    row++;

    const vals = keywords.map(function(kw) { return [kw, cat]; });
    rs.getRange(row, 1, vals.length, 2).setValues(vals).setBackground(rowColor).setFontSize(11);
    row += vals.length;
  });

  rs.getRange(row, 1, 1, 2).merge()
    .setValue("✍  CUSTOM & LEARNED RULES")
    .setBackground("#fef3c7").setFontColor("#92400e")
    .setFontWeight("bold").setFontSize(10).setVerticalAlignment("middle");
  rs.setRowHeight(row, 26);

  rs.setColumnWidth(1, 300);
  rs.setColumnWidth(2, 190);
  rs.hideSheet();
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
    out.push({ date: date, desc: cleanDesc(desc), amount: amount, source: String(fmt.name||"Custom").replace(/[^\w\s]/g, "").replace(/\s{2,}/g, " ").trim() });
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
  return HtmlService.createHtmlOutputFromFile('Modal').getContent();
}

// Compatibility wrapper.
function buildModalHTML() {
  return buildReviewImportModalHTML_();
}
