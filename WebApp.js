// ============================================================
// Budgets_WebApp.gs — Budget sheet setup + Web App entry point
// ============================================================

/**
 * Cria/reseta a aba Budgets com estrutura padrão.
 * Chamado pelo menu "💵 Setup Budget Sheet".
 */
function setupBudgetSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let bs = ss.getSheetByName(CONFIG.BUDGET_SHEET);
  if (!bs) bs = ss.insertSheet(CONFIG.BUDGET_SHEET);

  bs.clearContents();
  bs.clearFormats();

  // Row 1 — header
  bs.getRange(1,1,1,3).setValues([["Category","Monthly Budget ($)","Notes"]])
    .setBackground("#4f46e5").setFontColor("#fff").setFontWeight("bold");
  bs.setRowHeight(1, 32);

  // Row 2 — info banner
  bs.getRange(2,1,1,3).merge()
    .setValue("ℹ️   To hide this sheet: Budget Importer menu  →  💵 Budget")
    .setBackground("#e0f2fe").setFontColor("#0369a1").setFontStyle("italic")
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  bs.setRowHeight(2, 26);
  bs.setFrozenRows(2);

  // Categories from existing transactions, or defaults
  const expSheet = ss.getSheetByName(CONFIG.TARGET_SHEET);
  let cats = [];
  if (expSheet && expSheet.getLastRow() > 1) {
    const data = expSheet.getRange(2, CONFIG.COLS.CATEGORY, expSheet.getLastRow()-1, 1).getValues();
    const catSet = new Set();
    data.forEach(r => {
      const c = String(r[0]||"").trim();
      if (c && !CONFIG.INCOME_CATEGORIES.includes(c)) catSet.add(c);
    });
    cats = [...catSet].sort();
  }
  if (cats.length === 0) {
    cats = ["Appartament Rent","Groceries","Restaurant","Fast Food","GAS",
            "Mobile","Internet","GYM","Streamming","Coffee Shop",
            "Clothes","SkinCare","Car Payment","ICBC","Car Insurance"];
  }

  // Data starts at row 3
  const rows = cats.map(c => [c, 0, ""]);
  bs.getRange(3,1,rows.length,3).setValues(rows);
  bs.getRange(3,2,rows.length,1).setNumberFormat('"$"#,##0.00');

  bs.setColumnWidth(1,210); bs.setColumnWidth(2,170); bs.setColumnWidth(3,260);
}

function toggleBudgetSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let bs = ss.getSheetByName(CONFIG.BUDGET_SHEET);

  // Second click — hide
  if (bs && !bs.isSheetHidden()) { bs.hideSheet(); return; }

  // First click — ask to reset or just open
  const resp = ui.alert(
    "Budget",
    "Would you like to reset the Budget sheet?\n\n" +
    "YES — Rebuild with your latest transaction categories (current budget values will be lost).\n" +
    "NO  — Just open it as-is.",
    ui.ButtonSet.YES_NO
  );

  if (resp === ui.Button.YES) {
    setupBudgetSheet();
  } else if (!bs) {
    setupBudgetSheet();
  }

  bs = ss.getSheetByName(CONFIG.BUDGET_SHEET);
  if (!bs) return;
  if (bs.isSheetHidden()) bs.showSheet();
  ss.setActiveSheet(bs);
}

/**
 * Lê os budgets da aba Budgets.
 * Retorna: { "Groceries": 800, "Restaurant": 300, ... }
 */
function getBudgets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bs = ss.getSheetByName(CONFIG.BUDGET_SHEET);
  if (!bs || bs.getLastRow() < 2) return {};

  const data = bs.getRange(2, 1, bs.getLastRow() - 1, 2).getValues();
  const budgets = {};

  data.forEach(row => {
    const cat = String(row[0] || "").trim();
    const budget = parseAmountRobust(row[1]);
    if (cat && !cat.startsWith("ℹ") && budget > 0) budgets[cat] = budget;
  });

  return budgets;
}

/**
 * Salva budgets vindos do dashboard.
 * Recebe array de { cat, budget }.
 */
function saveBudgetsFromDashboard(budgetArray) {
  try {
    if (!Array.isArray(budgetArray)) {
      return { ok: false, error: "Invalid budget payload." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let bs = ss.getSheetByName(CONFIG.BUDGET_SHEET);
    if (!bs) {
      setupBudgetSheet();
      bs = ss.getSheetByName(CONFIG.BUDGET_SHEET);
    }

    const lastRow = bs.getLastRow();
    if (lastRow < 2) {
      return { ok: false, error: "Budget sheet has no rows. Run Setup Budget Sheet first." };
    }

    const data = bs.getRange(2, 1, lastRow - 1, 2).getValues();
    const rowByCategory = {};
    data.forEach((r, i) => {
      const cat = String(r[0] || "").trim();
      if (cat && !cat.startsWith("ℹ")) rowByCategory[cat] = i + 2;
    });

    budgetArray.forEach(item => {
      if (!item || !item.cat) return;
      const cat = String(item.cat).trim();
      const row = rowByCategory[cat];
      if (!row) return;
      const budget = Math.max(0, parseAmountRobust(item.budget));
      bs.getRange(row, 2).setValue(budget);
    });

    invalidateDashboardCache();
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

/**
 * Web App standalone.
 * Deploy: Execute as Me; access conforme sua necessidade.
 */
function doGet(e) {
  return HtmlService.createHtmlOutput(buildDashboardHTML())
    .setTitle("💰 Budget Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
