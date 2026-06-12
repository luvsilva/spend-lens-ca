// ============================================================
// Engine.gs — Cálculos puros + Cache
// ============================================================

const CACHE_KEY     = "dashboard_v3";
const CACHE_TTL_SEC = 21600; // 6 horas

function invalidateDashboardCache() {
  try { CacheService.getScriptCache().remove(CACHE_KEY); } catch(e) {}
}

function _saveToCache(data) {
  try {
    const json = JSON.stringify(data);
    if (json.length <= 95000) {
      CacheService.getScriptCache().put(CACHE_KEY, json, CACHE_TTL_SEC);
    } else {
      const lite = Object.assign({}, data);
      delete lite.txByCat;
      const liteJson = JSON.stringify(lite);
      if (liteJson.length <= 95000) {
        CacheService.getScriptCache().put(CACHE_KEY, liteJson, CACHE_TTL_SEC);
      } else {
        Logger.log("Dashboard cache skipped: payload too large (" + json.length + " bytes full, " + liteJson.length + " bytes lite)");
      }
    }
  } catch(e) {
    Logger.log("Dashboard cache write error: " + e.message);
  }
}

function _loadFromCache() {
  try {
    const json = CacheService.getScriptCache().get(CACHE_KEY);
    return json ? JSON.parse(json) : null;
  } catch(e) { return null; }
}

// ── Entry Point ───────────────────────────────────────────────
function getDashboardData(leakWindow, forceRefresh) {
  const reqWindow = leakWindow || 3;
  try {
    if (forceRefresh === true) {
      invalidateDashboardCache();
    }
    const cached = _loadFromCache();
    if (cached) {
      const budgetData = calcBudgetData(cached.currentMonthCats);
      let leaks = cached.leaks;
      // Only re-read the sheet when the requested leak window differs from what is cached.
      // This prevents a full sheet read on every cache hit.
      if (cached.leakWindow !== reqWindow) {
        const ss    = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET);
        if (sheet && sheet.getLastRow() > 1) {
          const tz     = Session.getScriptTimeZone();
          const raw    = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
          const parsed = _parseRows(raw, tz);
          leaks        = _calcLeaks(parsed, reqWindow);
          _saveToCache(Object.assign({}, cached, { leaks, leakWindow: reqWindow }));
        }
      }
      return {
        ok: true,
        data: Object.assign({}, cached, { budgetData, leaks }),
        fromCache: true
      };
    }

    // Cache miss: lê planilha uma única vez
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: null };

    const tz     = Session.getScriptTimeZone();
    const raw    = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
    const parsed = _parseRows(raw, tz);

    const analytics  = _calcAnalytics(parsed);
    const leaks      = _calcLeaks(parsed, reqWindow);
    const txByCat    = _buildTxByCat(parsed);
    const budgetData = calcBudgetData(analytics.currentMonthCats);
    const budgets    = getBudgets();

    const payload = {
      allMonthlyTotals: analytics.allMonthlyTotals,
      categoryTotals:   analytics.categoryTotals,
      monthCatData:     analytics.monthCatData,
      heatmapData:      analytics.heatmapData,
      availableMonths:  analytics.availableMonths,
      months:           analytics.months,
      currentMonthKey:  analytics.currentMonthKey,
      currentMonthCats: analytics.currentMonthCats,
      budgetData, leaks, budgets, txByCat,
      leakWindow: reqWindow
    };

    _saveToCache(payload);
    return { ok: true, data: payload, fromCache: false };

  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── getTxByCat ────────────────────────────────────────────────
function getTxByCat() {
  try {
    const cached = _loadFromCache();
    if (cached && cached.txByCat) return { ok: true, txByCat: cached.txByCat };

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { ok: true, txByCat: {} };
    const tz     = Session.getScriptTimeZone();
    const raw    = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
    const parsed = _parseRows(raw, tz);
    return { ok: true, txByCat: _buildTxByCat(parsed) };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── getTxForCategory ──────────────────────────────────────────
// Busca transações de UMA categoria diretamente da planilha.
// Mais rápido que carregar txByCat completo quando só precisa de uma cat.
function getTxForCategory(cat, monthKey, windowMonths) {
  try {
    if (!cat) return { ok: true, txs: [] };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { ok: true, txs: [] };
    const tz = Session.getScriptTimeZone();
    const raw = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    const parsed = _parseRows(raw, tz);
    let limitTs = 0;
    if (windowMonths && Number(windowMonths) > 0) {
      const months = Math.max(1, Math.min(12, Number(windowMonths)));
      const limit = new Date();
      limit.setMonth(limit.getMonth() - (months - 1));
      limit.setDate(1); limit.setHours(0, 0, 0, 0);
      limitTs = limit.getTime();
    }
    const txs = parsed
      .filter(function(r) {
        if (r.cat !== cat) return false;
        if (CONFIG.INCOME_CATEGORIES.includes(r.cat)) return false;
        if (monthKey && monthKey !== "all" && r.mk !== monthKey) return false;
        if (limitTs && r.ts < limitTs) return false;
        return true;
      })
      .sort(function(a, b) { return b.ts - a.ts; })
      .map(function(r) {
        return { date: r.dateStr, desc: r.desc, amount: Math.abs(r.amt), bank: r.bank, monthKey: r.mk };
      });
    return { ok: true, txs: txs };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ── Parse ─────────────────────────────────────────────────────
function _parseRows(raw, tz) {
  const result = [];
  raw.forEach(function(row) {
    const rawA = row[0], rawB = row[1];
    const desc = String(row[2]||"").trim();
    const cat  = String(row[3]||"").trim();
    const amt  = parseAmountRobust(row[4]);
    const bank = String(row[5]||"").trim();
    if (amt === 0 || !desc) return;
    let txDate = null;
    if (rawA instanceof Date && isValidDate(rawA))      txDate = rawA;
    else if (rawB instanceof Date && isValidDate(rawB)) txDate = rawB;
    else if (typeof rawB === "number" && rawB > 1000)
      txDate = new Date(Date.UTC(1899,11,30) + rawB*86400000);
    if (!txDate) {
      const s = String(rawA||"").trim().split(" ")[0].split("/");
      if (s.length===3 && s[2].length===4) txDate = new Date(+s[2],+s[1]-1,+s[0]);
    }
    if (!txDate || !isValidDate(txDate)) return;
    const mk      = Utilities.formatDate(txDate, tz, "yyyy-MM");
    const ml      = Utilities.formatDate(txDate, tz, "MMM yyyy");
    const dateStr = Utilities.formatDate(txDate, tz, "dd/MM/yyyy");
    result.push({ cat, desc, amt, bank, mk, ml, dateStr, ts: txDate.getTime() });
  });
  return result;
}

// ── Analytics ─────────────────────────────────────────────────
function _calcAnalytics(parsed) {
  const round = function(n){ return Math.round(n*100)/100; };
  const monthlyMap = {}, categoryMap = {}, monthCatMap = {};
  parsed.forEach(function(r) {
    if (CONFIG.INCOME_CATEGORIES.includes(r.cat) && r.amt > 0) return;
    const expense = Math.abs(r.amt);
    if (!monthlyMap[r.mk]) monthlyMap[r.mk] = { label: r.ml, total: 0 };
    monthlyMap[r.mk].total += expense;
    if (r.cat) {
      categoryMap[r.cat] = (categoryMap[r.cat]||0) + expense;
      if (!monthCatMap[r.mk]) monthCatMap[r.mk] = {};
      monthCatMap[r.mk][r.cat] = (monthCatMap[r.mk][r.cat]||0) + expense;
    }
  });
  const sortedKeys     = Object.keys(monthlyMap).sort();
  const categoryTotals = Object.entries(categoryMap)
    .sort(function(a,b){ return b[1]-a[1]; })
    .map(function(e){ return { cat:e[0], total:round(e[1]) }; });
  const monthCatData = {};
  Object.keys(monthCatMap).forEach(function(mk) {
    monthCatData[mk] = Object.entries(monthCatMap[mk])
      .sort(function(a,b){ return b[1]-a[1]; })
      .map(function(e){ return { cat:e[0], total:round(e[1]) }; });
  });
  const last12  = sortedKeys.slice(-12);
  const top10   = categoryTotals.slice(0,10).map(function(e){ return e.cat; });
  const curKey  = sortedKeys[sortedKeys.length-1] || "";
  const curCats = monthCatMap[curKey] || {};
  return {
    allMonthlyTotals: sortedKeys.map(function(k){
      return { key:k, label:monthlyMap[k].label, total:round(monthlyMap[k].total) };
    }),
    categoryTotals, monthCatData,
    heatmapData: {
      months: last12.map(function(k){ return monthlyMap[k].label; }),
      cats: top10,
      values: top10.map(function(cat){
        return last12.map(function(k){ return round((monthCatMap[k]||{})[cat]||0); });
      })
    },
    availableMonths: sortedKeys.map(function(k){ return { key:k, label:monthlyMap[k].label }; }),
    months: sortedKeys.length, currentMonthKey: curKey,
    currentMonthCats: Object.fromEntries(
      Object.entries(curCats).map(function(e){ return [e[0], round(e[1])]; })
    )
  };
}

// ── Leaks ─────────────────────────────────────────────────────
function _calcLeaks(parsed, windowMonths) {
  const months  = Math.max(1, Math.min(6, windowMonths));
  const limit   = new Date();
  limit.setMonth(limit.getMonth() - (months-1));
  limit.setDate(1); limit.setHours(0,0,0,0);
  const limitTs = limit.getTime();
  const catMap  = {};
  parsed.forEach(function(r) {
    if (!r.cat || r.amt===0 || CONFIG.INCOME_CATEGORIES.includes(r.cat) || r.ts<limitTs) return;
    const expense = Math.abs(r.amt);
    if (!catMap[r.cat]) catMap[r.cat] = { months: new Set(), total:0, count:0 };
    catMap[r.cat].months.add(r.mk);
    catMap[r.cat].total += expense;
    catMap[r.cat].count++;
  });
  const round = function(n){ return Math.round(n*100)/100; };
  const leaks = [];
  Object.entries(catMap).forEach(function(e) {
    const cat=e[0], d=e[1];
    const txCount=d.count, totalAmt=round(d.total), avgPerTx=totalAmt/txCount, monthCount=d.months.size;
    if (txCount>=4 && avgPerTx<=30 && totalAmt>=50 && monthCount>=Math.max(2,Math.ceil(months/2))) {
      leaks.push({ cat, txCount, totalAmt, monthCount,
        avgPerTx: round(avgPerTx), avgPerMonth: round(totalAmt/monthCount),
        freqPerWeek: round((txCount/(months*30))*7),
        level: totalAmt>=500?"high":totalAmt>=150?"medium":"low" });
    }
  });
  return leaks.sort(function(a,b){ return b.totalAmt-a.totalAmt; });
}

// ── TxByCat ───────────────────────────────────────────────────
function _buildTxByCat(parsed) {
  const result = {};
  parsed.forEach(function(r) {
    if (!r.cat || r.amt===0 || CONFIG.INCOME_CATEGORIES.includes(r.cat)) return;
    if (!result[r.cat]) result[r.cat] = [];
    result[r.cat].push({ date:r.dateStr, desc:r.desc, amount:Math.abs(r.amt), bank:r.bank, monthKey:r.mk, _ts:r.ts });
  });
  Object.keys(result).forEach(function(cat) {
    result[cat].sort(function(a,b){ return b._ts-a._ts; });
    result[cat].forEach(function(tx){ delete tx._ts; });
  });
  return result;
}

// ── Budget ────────────────────────────────────────────────────
function calcBudgetData(currentMonthCats) {
  const budgets = getBudgets();
  if (!budgets || Object.keys(budgets).length === 0) return null;

  const now = new Date();
  const lastDayOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  const daysPassed = now.getDate();
  const pctMonth = daysPassed / lastDayOfMonth;

  const items = Object.entries(budgets).map(function(e) {
    const cat = e[0], budget = e[1];
    const spent = Math.round((currentMonthCats[cat] || 0) * 100) / 100;
    const pct = budget > 0 ? Math.floor((spent / budget) * 100) : 0;
    const status = pct >= 100 ? "over" : pct >= 80 ? "warning" : "ok";

    let suggestion = null;
    if (pctMonth >= 0.85 && budget > 0 && spent > 0 && spent < budget) {
      const projected = (spent / daysPassed) * lastDayOfMonth;
      const rounded = Math.ceil(projected / 5) * 5;
      if (budget - rounded >= 5) suggestion = rounded;
    }

    return { cat, budget, spent, pct, status, suggestion };
  });

  items.sort(function(a, b) {
    const o = { over: 0, warning: 1, ok: 2 };
    return o[a.status] !== o[b.status] ? o[a.status] - o[b.status] : b.pct - a.pct;
  });

  return items;
}

// ── Compat: setupAnalysisSheet ────────────────────────────────
function calcAnalyticsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.TARGET_SHEET);
  if (!sh || sh.getLastRow()<2) return null;
  const tz = Session.getScriptTimeZone();
  return _calcAnalytics(_parseRows(sh.getRange(2,1,sh.getLastRow()-1,6).getValues(), tz));
}

function setupAnalysisSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const analytics = calcAnalyticsData();
  if (!analytics) { SpreadsheetApp.getUi().alert("No data yet."); return; }
  const { allMonthlyTotals, categoryTotals } = analytics;
  let an = ss.getSheetByName("Analysis");
  if (!an) an = ss.insertSheet("Analysis");
  else { an.clearContents(); an.clearFormats(); an.getCharts().forEach(c=>an.removeChart(c)); }
  const T1R=1,T1C=1;
  an.getRange(T1R,T1C,1,2).setValues([["Expense Date","sum Amount"]]);
  _styleHeader(an,T1R,T1C,1,2);
  if (allMonthlyTotals.length>0) {
    an.getRange(T1R+1,T1C,allMonthlyTotals.length,2).setValues(allMonthlyTotals.map(m=>[m.label,m.total]));
    an.getRange(T1R+1,T1C+1,allMonthlyTotals.length,1).setNumberFormat('"$"#,##0.00');
  }
  an.insertChart(an.newChart().setChartType(Charts.ChartType.COMBO)
    .addRange(an.getRange(T1R,T1C,allMonthlyTotals.length+1,2))
    .setOption("title","Monthly Expenses")
    .setOption("series",{0:{type:"bars",color:"#6366f1"},1:{type:"line",color:"#f87171",lineWidth:2,pointSize:4}})
    .setOption("legend",{position:"top"}).setOption("hAxis",{slantedText:true,slantedTextAngle:45})
    .setOption("vAxis",{format:"$#,##0"}).setOption("width",720).setOption("height",400)
    .setPosition(T1R,T1C+3,0,0).build());
  const T2R=T1R+allMonthlyTotals.length+3;
  an.getRange(T2R,T1C,1,2).setValues([["Category","SUM Amount"]]);
  _styleHeader(an,T2R,T1C,1,2);
  if (categoryTotals.length>0) {
    an.getRange(T2R+1,T1C,categoryTotals.length,2).setValues(categoryTotals.map(c=>[c.cat,c.total]));
    an.getRange(T2R+1,T1C+1,categoryTotals.length,1).setNumberFormat('"$"#,##0.00');
  }
  an.insertChart(an.newChart().setChartType(Charts.ChartType.BAR)
    .addRange(an.getRange(T2R,T1C,categoryTotals.length+1,2))
    .setOption("title","Expenses by Category").setOption("legend",{position:"none"})
    .setOption("hAxis",{format:"$#,##0"}).setOption("colors",["#6366f1"])
    .setOption("width",720).setOption("height",Math.max(400,categoryTotals.length*26+80))
    .setPosition(T2R,T1C+3,0,0).build());
  an.setColumnWidth(1,130); an.setColumnWidth(2,120);
  SpreadsheetApp.getUi().alert("✅ Analysis updated! "+allMonthlyTotals.length+" months.");
}

function _styleHeader(sheet,row,col,numRows,numCols) {
  const r=sheet.getRange(row,col,numRows,numCols);
  r.setBackground("#4f46e5"); r.setFontColor("#fff"); r.setFontWeight("bold");
}