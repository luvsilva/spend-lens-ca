// ============================================================
// Budgets_WebApp.gs — Budget sheet setup + Web App entry point
// ============================================================

// Spending layers — used to classify each category as Fixed / Variable / Lifestyle.
// This is only the DEFAULT seed; the user manages the real classification in the
// "Type" column of the Budgets sheet (dropdown). Unknown categories default to "Variable".
const BUDGET_LAYERS = ["Fixed", "Variable", "Lifestyle"];

const DEFAULT_BUDGET_TYPES_ = {
  // Fixed — recurring bills you don't act on month to month
  "Appartament Rent":"Fixed", "Car Payment":"Fixed", "ICBC":"Fixed",
  "Car Insurance":"Fixed", "Internet":"Fixed", "Mobile":"Fixed",
  "BCHydro":"Fixed", "Rent Insurance":"Fixed", "GYM":"Fixed",
  "Fees":"Fixed", "Banking Fees":"Fixed",
  // Variable — needs that move with behaviour
  "Groceries":"Variable", "GAS":"Variable", "Dog Food":"Variable",
  "Pharmacy":"Variable", "Health":"Variable", "Dental":"Variable",
  "Supplements":"Variable", "Car Maintenance":"Variable",
  "Public Transportation":"Variable", "Private Transportation":"Variable",
  "Parking":"Variable", "House Maintanence":"Variable", "SkinCare":"Variable",
  // Lifestyle — discretionary, the part to work on
  "Restaurant":"Lifestyle", "Fast Food":"Lifestyle", "Coffee Shop":"Lifestyle",
  "Streamming":"Lifestyle", "Clothes":"Lifestyle", "Leisure":"Lifestyle",
  "Night Club":"Lifestyle", "Vacation":"Lifestyle", "BarberShop":"Lifestyle",
  "BCLiquor":"Lifestyle"
};

function _defaultBudgetType_(cat) {
  return DEFAULT_BUDGET_TYPES_[cat] || "Variable";
}

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

  // Row 1 — header  (Category | Type | Monthly Budget | Notes)
  bs.getRange(1,1,1,4).setValues([["Category","Type","Monthly Budget ($)","Notes"]])
    .setBackground("#4f46e5").setFontColor("#fff").setFontWeight("bold");
  bs.setRowHeight(1, 32);

  // Row 2 — info banner
  bs.getRange(2,1,1,4).merge()
    .setValue("ℹ️   Type classifies each category as Fixed / Variable / Lifestyle.   To hide: Budget Importer menu → 💵 Budget")
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

  // Data starts at row 3  (Category | Type | Budget | Notes)
  const rows = cats.map(c => [c, _defaultBudgetType_(c), 0, ""]);
  bs.getRange(3,1,rows.length,4).setValues(rows);
  bs.getRange(3,3,rows.length,1).setNumberFormat('"$"#,##0.00');

  // Type dropdown (column 2)
  bs.getRange(3,2,Math.max(rows.length,1000),1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(BUDGET_LAYERS, true).build());

  bs.setColumnWidth(1,210); bs.setColumnWidth(2,110); bs.setColumnWidth(3,170); bs.setColumnWidth(4,260);
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

  // Layout: Category(1) | Type(2) | Budget(3) | Notes(4)
  const data = bs.getRange(2, 1, bs.getLastRow() - 1, 3).getValues();
  const budgets = {};

  data.forEach(row => {
    const cat = String(row[0] || "").trim();
    const budget = parseAmountRobust(row[2]);
    if (cat && !cat.startsWith("ℹ") && budget > 0) budgets[cat] = budget;
  });

  return budgets;
}

/**
 * Lê a coluna Type da aba Budgets.
 * Retorna: { "Groceries": "Variable", "Appartament Rent": "Fixed", ... }
 * Categorias sem Type definido caem no default (_defaultBudgetType_).
 */
function getBudgetTypes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bs = ss.getSheetByName(CONFIG.BUDGET_SHEET);
  if (!bs || bs.getLastRow() < 2) return {};

  const data = bs.getRange(2, 1, bs.getLastRow() - 1, 2).getValues();
  const types = {};

  data.forEach(row => {
    const cat = String(row[0] || "").trim();
    if (!cat || cat.startsWith("ℹ")) return;
    let type = String(row[1] || "").trim();
    if (BUDGET_LAYERS.indexOf(type) < 0) type = _defaultBudgetType_(cat);
    types[cat] = type;
  });

  return types;
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
      bs.getRange(row, 3).setValue(budget);  // Budget is column 3 (col 2 is Type)
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
