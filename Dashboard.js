// ============================================================
// Dashboard.gs — HTML do dashboard (desktop original preservado + mobile via JS detection)
// ============================================================

function openDashboard() {
  const html = HtmlService.createHtmlOutput(buildDashboardHTML()).setWidth(1400).setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, "💰 Budget Dashboard");
}

function buildDashboardHTML() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET);
  let cats = [];
  if (sheet && sheet.getLastRow() > 1) {
    const data = sheet.getRange(2, CONFIG.COLS.CATEGORY, sheet.getLastRow()-1, 1).getValues();
    const catSet = new Set();
    data.forEach(r => { const c = String(r[0]||"").trim(); if (c) catSet.add(c); });
    cats = [...catSet].sort();
  }
  const catsJson = JSON.stringify(cats);

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Budget">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/chartjs-plugin-datalabels/2.2.0/chartjs-plugin-datalabels.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');
:root {
  --radius:12px;
  --safe-top: env(safe-area-inset-top, 0px);
  --safe-bot: env(safe-area-inset-bottom, 0px);
}
[data-theme="dark"]{
  --bg:#0B1220;--surf:#111827;--surf2:#1F2937;--border:#1F2937;
  --accent:#818CF8;--red:#F87171;--green:#34D399;--yellow:#FBBF24;--purple:#C084FC;--orange:#FB923C;
  --text:#E5E7EB;--text2:#9CA3AF;--muted:#6B7280;--grid-c:rgba(31,41,55,.6);--tooltip:#1F2937;
}
[data-theme="light"]{
  --bg:#F8FAFC;--surf:#FFFFFF;--surf2:#F1F5F9;--border:#E2E8F0;
  --accent:#4F46E5;--red:#EF4444;--green:#22C55E;--yellow:#F59E0B;--purple:#A855F7;--orange:#F97316;
  --text:#0F172A;--text2:#475569;--muted:#64748B;--grid-c:rgba(226,232,240,.8);--tooltip:#1E293B;
}
*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent;}
html,body{width:100%;height:100%;overflow:hidden;font-family:'Inter',sans-serif;background:var(--bg);color:var(--text);font-size:12px;}
#app{display:flex;flex-direction:column;width:100%;height:100%;}

/* ══════════════════════════════════════════
   DESKTOP STYLES (original, untouched)
══════════════════════════════════════════ */
.topbar{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:8px 12px;background:var(--surf);border-bottom:1px solid var(--border);flex-shrink:0;overflow:hidden;}
.topbar-left{display:flex;align-items:center;gap:12px;flex:1 1 auto;min-width:0;overflow:hidden;}
.topbar-right{display:flex;align-items:center;gap:8px;flex:0 0 auto;min-width:max-content;}
.logo{font-size:14px;font-weight:800;flex:0 0 auto;max-width:280px;display:flex;align-items:center;gap:6px;white-space:nowrap;overflow:visible;}
.logo-brand{background:linear-gradient(135deg,var(--accent),var(--purple));-webkit-background-clip:text;-webkit-text-fill-color:transparent;}
.tabs{display:flex;gap:2px;background:var(--surf2);border:1px solid var(--border);border-radius:8px;padding:3px;}
.tab{padding:6px 12px;border-radius:6px;border:none;background:transparent;color:var(--text2);font-size:11px;font-weight:700;cursor:pointer;font-family:inherit;transition:all .15s;white-space:nowrap;}
.tab:hover{color:var(--text);}
.tab.active{background:var(--accent);color:#fff;}
.pills{display:flex;gap:3px;align-items:center;background:var(--surf2);border:1px solid var(--border);border-radius:8px;padding:3px 5px;flex:0 0 auto;min-width:max-content;max-width:none;overflow:visible;}
.pills label{font-size:10px;color:var(--muted);padding:0 5px;}
.pill{padding:4px 7px;border-radius:5px;border:none;background:transparent;color:var(--text2);font-size:10px;font-weight:700;cursor:pointer;font-family:inherit;transition:all .15s;white-space:nowrap;}
.pill:hover{background:var(--surf);color:var(--text);}
.pill.active{background:var(--accent);color:#fff;}
.income-wrap{display:flex;align-items:center;gap:4px;background:var(--surf2);border:1px solid var(--border);border-radius:8px;padding:5px 8px;flex:0 0 120px;min-width:110px;}
.income-wrap label{display:none;}
.income-input{background:transparent;border:none;color:var(--text);font-size:11px;font-weight:700;font-family:'JetBrains Mono',monospace;width:80px;outline:none;}
.income-input::placeholder{color:var(--muted);}
.btn{padding:5px 10px;border:none;border-radius:7px;font-size:11px;font-weight:700;cursor:pointer;font-family:inherit;transition:all .15s;white-space:nowrap;flex-shrink:0;}
.btn:hover{opacity:.85;}
.btn:active{transform:scale(.97);}
.btn-primary{background:var(--accent);color:#fff;}
.btn-ghost{background:var(--surf2);border:1px solid var(--border);color:var(--text2);}
.btn-ghost:hover{border-color:var(--accent);color:var(--accent);}
.theme-btn{display:flex;align-items:center;justify-content:center;gap:4px;background:var(--surf2);border:1px solid var(--border);border-radius:7px;padding:5px 8px;cursor:pointer;font-size:12px;font-weight:700;color:var(--text2);flex:0 0 34px;min-width:34px;flex-shrink:0;}
.theme-btn:hover{border-color:var(--accent);color:var(--accent);}
.kpi-row{display:flex;gap:7px;padding:8px 16px;flex-shrink:0;}
.kpi{flex:1;background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);padding:9px 12px;display:flex;align-items:center;gap:8px;}
.kpi:hover{border-color:var(--accent);}
.kpi-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0;}
.kpi .lbl{font-size:9px;color:var(--muted);font-weight:600;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px;}
.kpi .val{font-size:14px;font-weight:700;font-family:'JetBrains Mono',monospace;line-height:1;}
.c-red{color:var(--red);}.c-blue{color:var(--accent);}.c-green{color:var(--green);}.c-yellow{color:var(--yellow);}.c-purple{color:var(--purple);}
.panel{display:none;flex:1;overflow:hidden;min-height:0;flex-direction:column;}
.panel.active{display:flex;}
.grid{flex:1;display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);grid-template-rows:minmax(320px,1fr) minmax(320px,1fr);gap:14px;padding:0 18px 16px;overflow:auto;min-height:0;}
.cbox{background:var(--surf);border:1px solid var(--border);border-radius:18px;padding:14px 16px;display:flex;flex-direction:column;overflow:hidden;min-height:320px;box-shadow:0 10px 30px rgba(0,0,0,.10);}
.chead{display:flex;align-items:center;gap:8px;margin-bottom:12px;flex-shrink:0;}
.ctitle{font-size:13px;font-weight:800;text-transform:uppercase;letter-spacing:.7px;flex:1;}
.cwrap{flex:1;position:relative;min-height:260px;overflow:hidden;}
#chart-category-wrap{height:420px;max-height:420px;overflow-y:auto;overflow-x:hidden;padding-right:8px;scrollbar-gutter:stable;}
#chart-category-inner{position:relative;width:100%;min-height:420px;}
#chart-category-wrap::-webkit-scrollbar{width:6px;}
#chart-category-wrap::-webkit-scrollbar-thumb{background:var(--border);border-radius:4px;}
@media (max-width:1380px){
  .topbar{gap:8px;padding-left:10px;padding-right:10px;}
  .topbar-left{gap:8px;}
  .topbar-right{gap:6px;}
  .logo{max-width:260px;font-size:13px;}
  .tab{padding:6px 10px;}
  .pill{padding:4px 6px;font-size:9px;}
}
@media (max-width:1160px){
  .logo-brand{max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:inline-block;}
  .tabs{gap:1px;}
  .tab{padding:6px 8px;font-size:10px;}
  .pills label{display:none;}
  .pill{padding:4px 5px;}
  .income-wrap{flex-basis:96px;min-width:92px;padding-left:7px;padding-right:7px;}
  .income-input{width:62px;}
}

@media (max-width:1280px){
  .logo{max-width:250px;font-size:13px;}
  .tab{padding:6px 10px;font-size:10px;}
  .pill{padding:4px 6px;font-size:9px;}
  .pills label{display:none;}
  .income-wrap{flex-basis:105px;min-width:100px;}
  .income-input{width:70px;}
}

canvas{width:100%!important;height:100%!important;}
@media (min-width:1500px){.grid{grid-template-rows:minmax(390px,1fr) minmax(390px,1fr);gap:18px;}.cbox{min-height:390px;padding:18px 20px;}.cwrap{min-height:330px;}#chart-category-wrap{height:420px;max-height:420px;}}
@media (max-width:1050px){.grid{grid-template-columns:1fr;grid-template-rows:repeat(4,minmax(330px,auto));overflow:auto;}.cbox{min-height:330px;}}
.fsel{background:var(--surf2);border:1px solid var(--border);color:var(--text);border-radius:5px;padding:2px 7px;font-size:10px;font-family:inherit;cursor:pointer;outline:none;}
.hm-scroll{flex:1;overflow:auto;min-height:0;}
.hm-table{border-collapse:collapse;width:100%;}
.hm-table th{color:var(--muted);font-weight:600;padding:3px 5px;text-align:center;font-size:9px;white-space:nowrap;}
.hm-table td{padding:3px 5px;text-align:center;border-radius:3px;font-family:'JetBrains Mono',monospace;font-size:9px;white-space:nowrap;}
.hm-cat{text-align:right!important;color:var(--text);font-family:'Inter',sans-serif;font-weight:600;padding-right:8px!important;font-size:9px;}
.budget-panel{flex:1;overflow-y:auto;padding:0 16px 12px;min-height:0;}
.budget-header{display:flex;align-items:center;gap:10px;margin-bottom:12px;flex-shrink:0;padding:8px 16px 0;}
.budget-header h3{font-size:13px;font-weight:700;flex:1;}
.budget-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:8px;}
.bcard{background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);padding:12px 14px;cursor:pointer;transition:border-color .15s,transform .1s;}
.bcard:hover{border-color:var(--accent);transform:translateY(-1px);}
.bcard:active{transform:scale(.99);}
.bcard-top{display:flex;align-items:center;gap:8px;margin-bottom:8px;}
.bcat{font-size:12px;font-weight:600;flex:1;}
.bstatus{font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;}
.st-ok{background:rgba(52,211,153,.15);color:var(--green);}
.st-warning{background:rgba(251,191,36,.15);color:var(--yellow);}
.st-over{background:rgba(248,113,113,.15);color:var(--red);}
.st-none{background:var(--surf2);color:var(--muted);}
.bbar-bg{height:6px;background:var(--surf2);border-radius:3px;overflow:hidden;margin-bottom:6px;}
.bbar-fill{height:100%;border-radius:3px;transition:width .4s ease;}
.bcard-nums{display:flex;justify-content:space-between;font-size:10px;color:var(--text2);}
.bcard-nums .spent{font-weight:700;font-family:'JetBrains Mono',monospace;}
.bcard-edit{display:flex;align-items:center;gap:6px;margin-top:8px;padding-top:8px;border-top:1px solid var(--border);}
.bcard-edit label{font-size:10px;color:var(--muted);}
.budget-input{flex:1;background:var(--surf2);border:1px solid var(--border);color:var(--text);border-radius:5px;padding:3px 8px;font-size:11px;font-family:'JetBrains Mono',monospace;outline:none;}
.budget-input:focus{border-color:var(--accent);}
.bcard-hint{font-size:9px;color:var(--muted);margin-top:5px;}
.suggestion-bar{display:flex;align-items:center;gap:8px;margin-top:8px;padding:7px 10px;background:linear-gradient(135deg,rgba(99,102,241,.12),rgba(34,197,94,.1));border:1px solid rgba(99,102,241,.25);border-radius:8px;flex-wrap:wrap;}
.sugg-text{font-size:10px;color:var(--text2);flex:1;line-height:1.3;}
.sugg-btn{padding:4px 10px;background:var(--accent);color:#fff;border:none;border-radius:6px;font-size:10px;font-weight:700;cursor:pointer;white-space:nowrap;font-family:inherit;}
.sugg-btn:hover{opacity:.85;}
.remaining-card{background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);padding:14px 16px;margin:0 16px 8px;}
.rem-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--text2);margin-bottom:10px;}
.rem-row{display:flex;align-items:center;gap:12px;}
.rem-amount{font-size:24px;font-weight:700;font-family:'JetBrains Mono',monospace;}
.rem-bar-bg{flex:1;height:10px;background:var(--surf2);border-radius:5px;overflow:hidden;}
.rem-bar-fill{height:100%;border-radius:5px;transition:width .5s ease;}
.rem-labels{display:flex;justify-content:space-between;font-size:10px;color:var(--muted);margin-top:4px;}
.rem-projection{font-size:11px;color:var(--text2);margin-top:6px;}
/* ── Spending Layers ── */
.layers-wrap{margin:0 16px 10px;}
.layers-hero{background:linear-gradient(135deg,rgba(129,140,248,.16),rgba(192,132,252,.12));border:1px solid rgba(129,140,248,.3);border-radius:var(--radius);padding:14px 18px;margin-bottom:10px;}
.lh-label{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--text2);}
.lh-amount{font-size:30px;font-weight:800;font-family:'JetBrains Mono',monospace;color:var(--text);line-height:1.1;margin:2px 0;}
.lh-sub{font-size:11px;color:var(--text2);}
.layers-cards{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;}
.layer-card{background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);padding:12px 14px;cursor:pointer;transition:border-color .15s,transform .1s,opacity .15s;}
.layer-card:hover{transform:translateY(-1px);}
.layer-card.active{border-color:var(--accent);box-shadow:0 0 0 1px var(--accent);}
.layer-card.locked{opacity:.62;background:var(--surf2);}
.layer-card.locked:hover{opacity:.78;}
.lc-top{display:flex;align-items:center;gap:7px;margin-bottom:7px;}
.lc-icon{font-size:14px;}
.lc-name{font-size:12px;font-weight:700;flex:1;}
.lc-tag{font-size:8px;font-weight:800;letter-spacing:.4px;padding:2px 7px;border-radius:20px;background:var(--surf2);color:var(--muted);}
.layer-variable .lc-tag{background:rgba(251,146,60,.16);color:var(--orange);}
.layer-lifestyle .lc-tag{background:rgba(192,132,252,.16);color:var(--purple);}
.lc-amount{font-size:20px;font-weight:800;font-family:'JetBrains Mono',monospace;color:var(--text);}
.lc-bar-bg{height:6px;background:var(--surf2);border-radius:3px;overflow:hidden;margin:7px 0 5px;}
.lc-bar-fill{height:100%;border-radius:3px;transition:width .4s ease;}
.lc-sub{display:flex;justify-content:space-between;font-size:9px;color:var(--muted);}
body.is-mobile .layers-cards{grid-template-columns:1fr;}
.leaks-section{padding:0 16px 12px;flex:1;overflow-y:auto;}
.leak-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:8px;}
.leak-card{border-radius:var(--radius);padding:14px 16px;cursor:pointer;transition:transform .1s,box-shadow .15s;}
.leak-card:hover{transform:translateY(-2px);box-shadow:0 4px 20px rgba(0,0,0,.3);}
.cache-badge{font-size:9px;background:var(--green);color:#fff;padding:1px 6px;border-radius:10px;opacity:.7;}
#zoom-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);backdrop-filter:blur(4px);z-index:999;align-items:center;justify-content:center;}
#zoom-overlay.open{display:flex;}
#zoom-box{background:var(--surf);border:1px solid var(--border);border-radius:18px;padding:18px 20px;width:93vw;height:89vh;display:flex;flex-direction:column;box-shadow:0 25px 60px rgba(0,0,0,.4);}
#zoom-header{display:flex;align-items:center;margin-bottom:8px;}
#zoom-title{font-size:15px;font-weight:800;flex:1;}
#zoom-close{background:var(--surf2);border:1px solid var(--border);color:var(--text2);border-radius:7px;padding:5px 12px;cursor:pointer;font-family:inherit;font-size:11px;font-weight:600;}
#zoom-close:hover{background:var(--red);border-color:var(--red);color:#fff;}
#zoom-body{flex:1;position:relative;min-height:0;}
#zoom-cw{width:100%;height:100%;position:relative;overflow:auto;}
#zoom-canvas{
  width:100%!important;
  height:auto!important;
  min-height:420px;
}
#zoom-hm{flex:1;overflow:auto;height:100%;}
#drill-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);backdrop-filter:blur(3px);z-index:1000;align-items:center;justify-content:center;}
#drill-overlay.open{display:flex;}
#drill-box{background:var(--surf);border:1px solid var(--border);border-radius:16px;width:760px;max-width:96vw;max-height:84vh;display:flex;flex-direction:column;box-shadow:0 25px 60px rgba(0,0,0,.5);overflow:hidden;}
#drill-head{display:flex;align-items:center;gap:10px;padding:16px 20px 12px;border-bottom:1px solid var(--border);flex-shrink:0;}
#drill-icon{font-size:24px;}
#drill-info{flex:1;min-width:0;}
#drill-name{font-size:15px;font-weight:700;}
#drill-sub{font-size:10px;color:var(--muted);margin-top:1px;}
#drill-close{background:var(--surf2);border:1px solid var(--border);color:var(--text2);border-radius:7px;padding:5px 12px;cursor:pointer;font-family:inherit;font-size:11px;font-weight:600;flex-shrink:0;}
#drill-close:hover{background:var(--red);border-color:var(--red);color:#fff;}
#drill-stats{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;padding:10px 20px;border-bottom:1px solid var(--border);flex-shrink:0;background:var(--surf2);}
.dstat{text-align:center;}
.dstat-lbl{font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:.4px;margin-bottom:2px;}
.dstat-val{font-size:14px;font-weight:700;font-family:'JetBrains Mono',monospace;}
#drill-body{flex:1;overflow-y:auto;min-height:0;-webkit-overflow-scrolling:touch;}
#drill-loading{padding:40px;text-align:center;color:var(--muted);}
#drill-table{width:100%;border-collapse:collapse;display:none;}
#drill-table th{font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.4px;padding:8px 16px;text-align:left;border-bottom:1px solid var(--border);position:sticky;top:0;background:var(--surf);}
#drill-table th.th-amt,#drill-table th.th-recat{text-align:center;}
#drill-table td{font-size:11px;padding:8px 16px;border-bottom:1px solid rgba(46,50,80,.3);}
#drill-table tr:hover td{background:var(--surf2);}
.td-date{font-family:'JetBrains Mono',monospace;color:var(--muted);white-space:nowrap;width:88px;}
.td-desc{color:var(--text);max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.td-amt{font-family:'JetBrains Mono',monospace;font-weight:700;color:var(--red);text-align:right;white-space:nowrap;width:80px;}
.td-bank{font-size:10px;color:var(--muted);white-space:nowrap;width:54px;}
.td-recat{width:140px;text-align:center;}
.recat-sel{background:var(--surf2);border:1px solid var(--border);color:var(--text);border-radius:5px;padding:3px 6px;font-size:10px;font-family:inherit;cursor:pointer;outline:none;width:130px;}
.recat-sel:hover{border-color:var(--accent);}
.recat-sel.saving{opacity:.5;pointer-events:none;}
.recat-sel.saved{border-color:var(--green);}
::-webkit-scrollbar{width:4px;height:4px;}
::-webkit-scrollbar-thumb{background:var(--border);border-radius:2px;}
#loading{text-align:center;padding:80px;color:var(--muted);font-size:14px;}
#error-msg{display:none;color:var(--red);text-align:center;padding:60px;font-weight:600;}
#app-inner{display:none;flex:1;overflow:hidden;flex-direction:column;min-height:0;}
#phase2-bar{height:2px;background:linear-gradient(90deg,var(--accent),var(--purple));width:0%;transition:width .5s;position:fixed;top:0;left:0;z-index:9999;display:none;}

/* ══════════════════════════════════════════
   MOBILE-ONLY ELEMENTS — hidden by default
   Shown only via JS class .is-mobile on body
══════════════════════════════════════════ */

/* Mobile bottom nav */
.mobile-nav{
  display:none;
  position:fixed;bottom:0;left:0;right:0;
  background:var(--surf);border-top:1px solid var(--border);
  padding:8px 0;padding-bottom:calc(8px + env(safe-area-inset-bottom, 0px));
  z-index:500;flex-shrink:0;
}
.mobile-nav-inner{display:flex;justify-content:space-around;align-items:center;}
.mobile-tab{display:flex;flex-direction:column;align-items:center;gap:3px;padding:6px 20px;border:none;background:transparent;cursor:pointer;color:var(--muted);font-family:inherit;transition:color .15s;}
.mobile-tab.active{color:var(--accent);}
.mobile-tab-icon{font-size:20px;line-height:1;}
.mobile-tab-label{font-size:9px;font-weight:600;text-transform:uppercase;letter-spacing:.4px;}

/* Mobile extra rows */
.mobile-topbar-row2{display:none;padding:6px 12px 8px;gap:8px;background:var(--surf);border-bottom:1px solid var(--border);flex-shrink:0;overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none;}
.mobile-topbar-row2::-webkit-scrollbar{display:none;}
.mobile-pills-row{display:none;padding:6px 12px;background:var(--surf2);border-bottom:1px solid var(--border);flex-shrink:0;overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none;gap:4px;align-items:center;}
.mobile-pills-row::-webkit-scrollbar{display:none;}

/* Mobile chart cards */
.mobile-chart-scroll{flex:1;overflow-y:auto;-webkit-overflow-scrolling:touch;padding:8px 12px 12px;}
.mobile-cbox{background:var(--surf);border:1px solid var(--border);border-radius:var(--radius);padding:10px 12px;margin-bottom:10px;}
.mobile-chart-wrap{position:relative;height:280px;}
body.is-mobile #chart-category-wrap-m{height:430px;max-height:430px;overflow-y:auto;overflow-x:hidden;padding-right:4px;scrollbar-gutter:stable;}
#chart-category-inner-m{position:relative;width:100%;min-height:300px;}
#chart-category-wrap-m::-webkit-scrollbar{width:5px;}
#chart-category-wrap-m::-webkit-scrollbar-thumb{background:var(--border);border-radius:4px;}
body.is-mobile .mobile-cbox{border-radius:18px;padding:14px 14px;margin-bottom:14px;}

/* ── .is-mobile applied by JS when screen <= 600px ── */
body.is-mobile #app{height:100dvh;overflow:hidden;}
body.is-mobile .topbar{padding:10px 14px;padding-top:calc(10px + env(safe-area-inset-top,0px));gap:8px;}
body.is-mobile .topbar-left{flex:1;gap:8px;}
body.is-mobile .topbar-right{display:flex;}
body.is-mobile .tabs{display:none;}
body.is-mobile .pills{display:none !important;}
body.is-mobile .topbar > .topbar-right .income-wrap{display:none;}
body.is-mobile .topbar > .topbar-right .btn-primary{display:none;}
body.is-mobile .topbar > .topbar-right .theme-btn{display:none;}
body.is-mobile .btn-close-modal{display:none;}
body.is-mobile .logo{font-size:13px;}
body.is-mobile .mobile-nav{display:block;}
body.is-mobile .mobile-topbar-row2{display:flex;}
body.is-mobile .mobile-pills-row{display:flex;}
body.is-mobile .kpi-row{padding:8px 12px;gap:6px;overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none;flex-wrap:nowrap;}
body.is-mobile .kpi-row::-webkit-scrollbar{display:none;}
body.is-mobile .kpi{min-width:130px;padding:8px 10px;flex:0 0 auto;}
body.is-mobile .kpi .val{font-size:13px;}
body.is-mobile .panel.active{display:flex;flex-direction:column;overflow:hidden;}
body.is-mobile #panel-charts .grid{display:none;}
body.is-mobile #mobile-charts-inner{display:flex !important;flex-direction:column;flex:1;overflow:hidden;min-height:0;}
body.is-mobile .budget-header{padding:10px 12px 0;}
body.is-mobile .budget-header h3{font-size:12px;}
body.is-mobile .budget-panel{padding:0 12px calc(80px + env(safe-area-inset-bottom,0px));}
body.is-mobile .budget-grid{grid-template-columns:1fr;gap:8px;}
body.is-mobile .bcard{padding:10px 12px;}
body.is-mobile .bcat{font-size:11px;}
body.is-mobile .leaks-section{padding:0 12px calc(80px + env(safe-area-inset-bottom,0px));}
body.is-mobile .leak-grid{grid-template-columns:1fr;}
body.is-mobile .remaining-card{margin:8px 12px 6px;padding:12px;}
body.is-mobile .rem-amount{font-size:20px;}
body.is-mobile #drill-box{width:100vw;max-width:100vw;max-height:95vh;border-radius:16px 16px 0 0;align-self:flex-end;}
body.is-mobile #drill-overlay{align-items:flex-end;}
body.is-mobile #drill-head{padding:14px 16px 10px;}
body.is-mobile #drill-name{font-size:13px;}
body.is-mobile #drill-stats{grid-template-columns:repeat(2,1fr);padding:8px 16px;gap:6px;}
body.is-mobile .dstat-val{font-size:12px;}
body.is-mobile #drill-table th,body.is-mobile #drill-table td{padding:8px 10px;font-size:10px;}
body.is-mobile .td-desc{max-width:140px;}
body.is-mobile .recat-sel{width:110px;font-size:9px;}
body.is-mobile #zoom-box{width:100vw;height:100vh;border-radius:0;padding:16px;}

/* Fix: avoid duplicated month filter after switching tabs */
body.hide-months #pills-wrap,
body.hide-months #mobile-pills-row{display:none!important;}

/* Zoom header controls */
#zoom-controls{display:flex;align-items:center;gap:8px;margin-right:10px;}
#zoom-controls .fsel{font-size:11px;padding:5px 10px;border-radius:8px;}
body.is-mobile #zoom-header{gap:8px;}
body.is-mobile #zoom-title{font-size:13px;}
body.is-mobile #zoom-controls{margin-right:0;}
body.is-mobile #zoom-controls .fsel{max-width:140px;}

</style>
</head>
<body data-theme="dark">
<div id="phase2-bar"></div>
<div id="app">

  <!-- ══ Top Bar ══ -->
  <div class="topbar">
    <div class="topbar-left">
      <div class="logo"><span>💰</span><span class="logo-brand">Budget Dashboard</span><span id="cache-badge" class="cache-badge" style="display:none">⚡ cached</span></div>
      <div class="tabs">
        <button class="tab active" onclick="showPanel('charts')">📊 Charts</button>
        <button class="tab" onclick="showPanel('budget')">💵 Budget</button>
        <button class="tab" onclick="showPanel('leaks')">💧 Leaks</button>
      </div>
      <div class="pills" id="pills-wrap">
        <label>Months:</label>
        <button class="pill" onclick="setPill(1)">1</button><button class="pill" onclick="setPill(2)">2</button><button class="pill" onclick="setPill(3)">3</button><button class="pill" onclick="setPill(4)">4</button><button class="pill" onclick="setPill(5)">5</button><button class="pill" onclick="setPill(6)">6</button><button class="pill" onclick="setPill(7)">7</button><button class="pill" onclick="setPill(8)">8</button><button class="pill" onclick="setPill(9)">9</button><button class="pill active" onclick="setPill(10)">10</button><button class="pill" onclick="setPill(11)">11</button><button class="pill" onclick="setPill(12)">12</button><button class="pill" onclick="setPill(999)">All</button>
      </div>
    </div>
    <div class="topbar-right">
      <div class="income-wrap">
        <label>Monthly Income $</label>
        <input class="income-input" id="income-input" type="number" placeholder="Income" onchange="updateRemaining()"/>
      </div>
      <div class="theme-btn" onclick="toggleTheme()" id="theme-btn" title="Toggle theme">☀️</div>
      <button class="btn btn-primary" onclick="loadData(true)" title="Refresh">🔄</button>
      <button class="btn btn-ghost btn-close-modal" onclick="safeCloseDashboard()" title="Close">✕</button>
    </div>
  </div>

  <!-- ══ Mobile Row 2: income + theme (hidden on desktop) ══ -->
  <div class="mobile-topbar-row2">
    <div class="income-wrap" style="flex:1">
      <label>Income $</label>
      <input class="income-input" id="income-input-mobile" type="number" placeholder="Monthly income" oninput="syncIncome(this)"/>
    </div>
    <div class="theme-btn" onclick="toggleTheme()" style="flex-shrink:0">☀️</div>
    <button class="btn btn-primary" onclick="loadData(true)" style="flex-shrink:0">🔄</button>
  </div>

  <!-- ══ Mobile Pills Row (hidden on desktop) ══ -->
  <div class="mobile-pills-row" id="mobile-pills-row">
    <span style="font-size:10px;color:var(--muted);font-weight:600;white-space:nowrap;flex-shrink:0">Months:</span>
    <button class="pill" onclick="setPill(1)">1</button><button class="pill" onclick="setPill(2)">2</button><button class="pill" onclick="setPill(3)">3</button><button class="pill" onclick="setPill(4)">4</button><button class="pill" onclick="setPill(5)">5</button><button class="pill" onclick="setPill(6)">6</button><button class="pill" onclick="setPill(7)">7</button><button class="pill" onclick="setPill(8)">8</button><button class="pill" onclick="setPill(9)">9</button><button class="pill active" onclick="setPill(10)">10</button><button class="pill" onclick="setPill(11)">11</button><button class="pill" onclick="setPill(12)">12</button><button class="pill" onclick="setPill(999)">All</button>
  </div>

  <div id="loading">⏳ Loading data...</div>
  <div id="error-msg"></div>

  <div id="app-inner">
    <!-- KPI Row -->
    <div class="kpi-row">
      <div class="kpi"><div class="kpi-dot" style="background:var(--red)"></div><div><div class="lbl">Total Spent</div><div class="val c-red" id="kpi-total">—</div></div></div>
      <div class="kpi"><div class="kpi-dot" style="background:var(--accent)"></div><div><div class="lbl">Months</div><div class="val c-blue" id="kpi-months">—</div></div></div>
      <div class="kpi"><div class="kpi-dot" style="background:var(--green)"></div><div><div class="lbl">Avg/Month</div><div class="val c-green" id="kpi-avg">—</div></div></div>
      <div class="kpi"><div class="kpi-dot" style="background:var(--yellow)"></div><div><div class="lbl">This Month</div><div class="val c-yellow" id="kpi-current">—</div></div></div>
      <div class="kpi"><div class="kpi-dot" style="background:var(--green)"></div><div><div class="lbl">Best Month</div><div class="val c-green" id="kpi-best">—</div></div></div>
      <div class="kpi"><div class="kpi-dot" style="background:var(--red)"></div><div><div class="lbl">Worst Month</div><div class="val c-red" id="kpi-worst">—</div></div></div>
      <div class="kpi"><div class="kpi-dot" style="background:var(--purple)"></div><div><div class="lbl">Remaining</div><div class="val c-purple" id="kpi-remaining">—</div></div></div>
    </div>

    <!-- ══ CHARTS PANEL ══ -->
    <div class="panel active" id="panel-charts">
      <!-- Desktop 2x2 grid -->
      <div class="grid">
        <div class="cbox"><div class="chead"><div class="ctitle">📅 Monthly Expenses</div><button class="btn btn-ghost" style="padding:2px 7px;font-size:10px" onclick="openZoom('monthly')">⛶</button></div><div class="cwrap"><canvas id="chart-monthly"></canvas></div></div>
        <div class="cbox"><div class="chead"><div class="ctitle">🏷️ Top 10 Categories</div><select id="sel-cat-month" class="fsel" onchange="updateCategoryChart()"><option value="all">All time</option></select><button class="btn btn-ghost" style="padding:2px 7px;font-size:10px;margin-left:4px" onclick="openZoom('category')">⛶</button></div><div class="cwrap" id="chart-category-wrap"><div id="chart-category-inner"><canvas id="chart-category"></canvas></div></div></div>
        <div class="cbox"><div class="chead"><div class="ctitle">🔥 Category Heatmap</div><button class="btn btn-ghost" style="padding:2px 7px;font-size:10px;margin-left:auto" onclick="openZoom('heatmap')">⛶</button></div><div class="hm-scroll" id="hm-main"></div></div>
        <div class="cbox"><div class="chead"><div class="ctitle">📈 Spending Trend</div><button class="btn btn-ghost" style="padding:2px 7px;font-size:10px;margin-left:auto" onclick="openZoom('trend')">⛶</button></div><div class="cwrap"><canvas id="chart-trend"></canvas></div></div>
      </div>
      <!-- Mobile scrollable charts (shown only via .is-mobile) -->
      <div id="mobile-charts-inner" style="display:none;flex-direction:column;flex:1;overflow:hidden;min-height:0;">
        <div class="mobile-chart-scroll">
          <div class="mobile-cbox">
            <div class="chead"><div class="ctitle">📅 Monthly Expenses</div><button class="btn btn-ghost" style="padding:2px 7px;font-size:10px" onclick="openZoom('monthly')">⛶</button></div>
            <div class="mobile-chart-wrap"><canvas id="chart-monthly-m"></canvas></div>
          </div>
          <div class="mobile-cbox">
            <div class="chead"><div class="ctitle">🏷️ Top 10 Categories</div><select id="sel-cat-month-m" class="fsel" onchange="updateCategoryChartMobile()"><option value="all">All time</option></select></div>
            <div class="mobile-chart-wrap" id="chart-category-wrap-m" style="height:430px"><div id="chart-category-inner-m"><canvas id="chart-category-m"></canvas></div></div>
          </div>
          <div class="mobile-cbox">
            <div class="chead"><div class="ctitle">📈 Spending Trend</div></div>
            <div class="mobile-chart-wrap"><canvas id="chart-trend-m"></canvas></div>
          </div>
          <div class="mobile-cbox">
            <div class="chead"><div class="ctitle">🔥 Category Heatmap</div></div>
            <div class="hm-scroll" id="hm-mobile" style="max-height:300px"></div>
          </div>
        </div>
      </div>
    </div>

    <!-- ══ BUDGET PANEL ══ -->
    <div class="panel" id="panel-budget">
      <div class="remaining-card" id="remaining-card" style="display:none">
        <div class="rem-title">💵 Monthly Remaining</div>
        <div class="rem-row">
          <div class="rem-amount" id="rem-amount">—</div>
          <div style="flex:1"><div class="rem-bar-bg"><div class="rem-bar-fill" id="rem-bar" style="width:0%"></div></div><div class="rem-labels"><span id="rem-spent-lbl">Spent: $0</span><span id="rem-income-lbl">Income: —</span></div></div>
        </div>
        <div class="rem-projection" id="rem-projection"></div>
      </div>

      <!-- ══ SPENDING LAYERS ══ -->
      <div class="layers-wrap" id="layers-wrap" style="display:none">
        <div class="layers-hero">
          <div class="lh-label">🎯 Controllable this month</div>
          <div class="lh-amount" id="lh-controllable">—</div>
          <div class="lh-sub">Variable + Lifestyle — the part you can actually work on (Fixed is locked)</div>
        </div>
        <div class="layers-cards" id="layers-cards"></div>
      </div>

      <div class="budget-header">
        <h3>💵 Budget by Category <span style="font-size:10px;color:var(--muted);font-weight:400">(current month · click card to view transactions)</span><span id="budget-filter-lbl" style="font-size:10px;color:var(--accent);font-weight:700"></span></h3>
        <button class="btn btn-primary" onclick="saveBudgets()" id="save-budget-btn">💾 Save Budgets</button>
      </div>
      <div class="budget-panel"><div class="budget-grid" id="budget-grid"></div></div>
    </div>

    <!-- ══ LEAKS PANEL ══ -->
    <div class="panel" id="panel-leaks">
      <div class="leaks-section">
        <div style="display:flex;align-items:center;gap:12px;margin-bottom:14px;margin-top:8px;">
          <div style="font-size:13px;font-weight:700">💧 Spending Leaks</div>
          <div style="font-size:11px;color:var(--muted)">Small but frequent charges that add up · click card to inspect</div>
          <div style="margin-left:auto;display:flex;align-items:center;gap:6px;background:var(--surf2);border:1px solid var(--border);border-radius:8px;padding:4px 10px;">
            <label style="font-size:10px;color:var(--muted);font-weight:600">Window:</label>
            <select id="leak-window" class="fsel" onchange="reloadLeaks()" style="border:none;background:transparent;padding:0">
              <option value="1">1 month</option><option value="2">2 months</option><option value="3" selected>3 months</option><option value="4">4 months</option><option value="5">5 months</option><option value="6">6 months</option>
            </select>
          </div>
        </div>
        <div style="font-size:10px;color:var(--muted);margin-bottom:12px;background:var(--surf2);border:1px solid var(--border);border-radius:8px;padding:8px 12px;display:flex;gap:16px;flex-wrap:wrap;">
          <span>🔍 <strong style="color:var(--text2)">≥4 transactions</strong></span>
          <span>💰 <strong style="color:var(--text2)">≤$30 per charge</strong></span>
          <span>📊 <strong style="color:var(--text2)">≥$50 total</strong></span>
          <span>📅 <strong style="color:var(--text2)">Recurring across months</strong></span>
        </div>
        <div class="leak-grid" id="leak-grid"></div>
      </div>
    </div>
  </div><!-- end app-inner -->

  <!-- ══ Mobile Bottom Nav (hidden on desktop) ══ -->
  <nav class="mobile-nav">
    <div class="mobile-nav-inner">
      <button class="mobile-tab active" id="mnav-charts" onclick="mobileShowPanel('charts')">
        <span class="mobile-tab-icon">📊</span><span class="mobile-tab-label">Charts</span>
      </button>
      <button class="mobile-tab" id="mnav-budget" onclick="mobileShowPanel('budget')">
        <span class="mobile-tab-icon">💵</span><span class="mobile-tab-label">Budget</span>
      </button>
      <button class="mobile-tab" id="mnav-leaks" onclick="mobileShowPanel('leaks')">
        <span class="mobile-tab-icon">💧</span><span class="mobile-tab-label">Leaks</span>
      </button>
    </div>
  </nav>
</div><!-- end app -->

<!-- ══ ZOOM OVERLAY ══ -->
<div id="zoom-overlay" onclick="closeZoomOutside(event)">
  <div id="zoom-box">
    <div id="zoom-header"><div id="zoom-title">Chart</div><div id="zoom-controls"></div><button id="zoom-close" onclick="closeZoom()">✕ Close</button></div>
    <div id="zoom-body">
      <div id="zoom-cw" style="display:none"><canvas id="zoom-canvas"></canvas></div>
      <div id="zoom-hm" style="display:none"></div>
    </div>
  </div>
</div>

<!-- ══ DRILL-DOWN OVERLAY ══ -->
<div id="drill-overlay">
  <div id="drill-box">
    <div id="drill-head">
      <span id="drill-icon">📋</span>
      <div id="drill-info"><div id="drill-name">Category</div><div id="drill-sub">Loading...</div></div>
      <button id="drill-close" onclick="closeDrill()">✕ Close</button>
    </div>
    <div id="drill-stats">
      <div class="dstat"><div class="dstat-lbl">Transactions</div><div class="dstat-val c-blue" id="ds-count">—</div></div>
      <div class="dstat"><div class="dstat-lbl">Total Spent</div><div class="dstat-val c-red" id="ds-total">—</div></div>
      <div class="dstat"><div class="dstat-lbl">Avg/charge</div><div class="dstat-val c-yellow" id="ds-avg">—</div></div>
      <div class="dstat"><div class="dstat-lbl">Largest</div><div class="dstat-val c-purple" id="ds-max">—</div></div>
    </div>
    <div id="drill-body">
      <div id="drill-loading">⏳ Loading transactions...</div>
      <table id="drill-table">
        <thead><tr>
          <th>Date</th><th>Description</th>
          <th class="th-amt">Amount</th><th>Bank</th>
          <th class="th-recat">Reclassify</th>
        </tr></thead>
        <tbody id="drill-tbody"></tbody>
      </table>
    </div>
  </div>
</div>

<script>
Chart.register(ChartDataLabels);

// ── Mobile detection via JS (not media queries) ───────────────
var IS_MOBILE = typeof window !== 'undefined' && (window.innerWidth <= 700 || (window.screen && window.screen.width <= 600));
function initMobileClass(){
  if(IS_MOBILE){ document.body.classList.add('is-mobile'); }
}

// ── Theme ─────────────────────────────────────────────────────
const THEMES={
  dark:{bg:"#0B1220",surface:"#111827",income:"#34D399",expense:"#F87171",warning:"#FBBF24",trend:"#34D399",avg:"#FBBF24",cats:["#818CF8","#22D3EE","#C084FC","#FB923C","#2DD4BF","#A3E635","#F87171","#34D399","#FBBF24","#60A5FA"],text:"#E5E7EB",muted:"#9CA3AF",border:"#1F2937",grid:"rgba(31,41,55,.6)",tooltip:"#1F2937"},
  light:{bg:"#F8FAFC",surface:"#FFFFFF",income:"#22C55E",expense:"#EF4444",warning:"#F59E0B",trend:"#4F46E5",avg:"#F59E0B",cats:["#4F46E5","#06B6D4","#A855F7","#F97316","#14B8A6","#84CC16","#EF4444","#F59E0B","#22C55E","#8B5CF6"],text:"#0F172A",muted:"#64748B",border:"#E2E8F0",grid:"rgba(226,232,240,.8)",tooltip:"#1E293B"}
};
function TH(){return THEMES[currentTheme]||THEMES.dark;}
function P(){return TH().cats;}

const CAT_ICONS={"Coffee Shop":"☕","Restaurant":"🍽️","Fast Food":"🍔","Groceries":"🛒","GAS":"⛽","Parking":"🅿️","Streamming":"📺","GYM":"🏋️","Supplements":"💊","Clothes":"👕","SkinCare":"🧴","Leisure":"🎭","Mobile":"📱","Internet":"🌐","Pharmacy":"💊","Banking Fees":"🏦","Fees":"🏦","Public Transportation":"🚌","BCHydro":"⚡","BCLiquor":"🍷","Car Payment":"🚗","Car Maintenance":"🔧","Appartament Rent":"🏠","Vacation":"✈️","Health":"❤️","Dental":"🦷","BarberShop":"💈","Dog Food":"🐾","House Maintanence":"🏠","Courses":"📚","Tax Return":"💰","GIFTS":"🎁","Night Club":"🎉","Eletronics":"💻","Gadgets":"🔌","ICBC":"🛡️","Car Insurance":"🛡️","Rent Insurance":"📋","Cash Withdrawal":"💵"};
const ALL_CATS=${catsJson};

let allData=null,charts={},zoomChart=null,activePills=10,currentTheme="dark";
let txByCatReady=false;

const fmt=n=>"$"+Number(n).toLocaleString("en-CA",{minimumFractionDigits:0,maximumFractionDigits:0});
const fmtFull=n=>"$"+Number(n).toLocaleString("en-CA",{minimumFractionDigits:2,maximumFractionDigits:2});
const esc=v=>String(v==null?'':v).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c]);
const cssVar=n=>getComputedStyle(document.documentElement).getPropertyValue(n).trim();
function destroyAll(){Object.values(charts).forEach(c=>{try{c.destroy();}catch(e){}});charts={};}

function toggleTheme(){
  currentTheme=currentTheme==="dark"?"light":"dark";
  document.body.setAttribute("data-theme",currentTheme);
  var label=currentTheme==="dark"?"☀️":"🌙";
  document.getElementById("theme-btn").textContent=label;
  document.querySelectorAll(".mobile-topbar-row2 .theme-btn").forEach(b=>{b.textContent=currentTheme==="dark"?"☀️":"🌙";});
  if(allData){destroyAll();renderAllCharts();}
}

// ── Panel switching ───────────────────────────────────────────
function showPanel(name){
  document.querySelectorAll(".panel").forEach(p=>p.classList.remove("active"));
  document.querySelectorAll(".tab").forEach(t=>t.classList.remove("active"));
  document.getElementById("panel-"+name).classList.add("active");
  document.querySelectorAll(".tab").forEach(t=>{if(t.textContent.toLowerCase().includes(name.substring(0,4)))t.classList.add("active");});

  // Fix: do not manually toggle desktop/mobile month rows.
  // CSS controls visibility, preventing duplicate month filters after navigation.
  document.body.classList.toggle("hide-months", name !== "charts");

  // Warm transaction cache when user enters detail-heavy tabs.
  if(name === "budget" || name === "leaks") warmTxCacheForCurrentViews();
}
function mobileShowPanel(name){
  showPanel(name);
  document.querySelectorAll(".mobile-tab").forEach(t=>t.classList.remove("active"));
  var nb=document.getElementById("mnav-"+name);
  if(nb)nb.classList.add("active");
}

// ── Pills ─────────────────────────────────────────────────────
function setPill(n){
  activePills = n;
  const labels = [1,2,3,4,5,6,7,8,9,10,11,12,999];
  document.querySelectorAll(".pill").forEach((p, i) => {
    p.classList.toggle("active", labels[i % labels.length] === n);
  });
  if (!allData) return;
  destroyAll();
  renderAllCharts();
}

// ── Chart configs ─────────────────────────────────────────────
function monthlyConfig(data,isMobile){
  const th=TH();
  return{data:{labels:data.map(m=>m.label),datasets:[
    {type:"bar",label:"Spent",data:data.map(m=>m.total),
      backgroundColor:data.map((_,i)=>i===data.length-1?th.warning+"cc":th.expense+"66"),
      borderColor:data.map((_,i)=>i===data.length-1?th.warning:th.expense),
      borderWidth:1.5,borderRadius:6,yAxisID:"y",
      datalabels:{anchor:"end",align:"end",color:th.text,font:{size:isMobile?8:9,weight:"700",family:"JetBrains Mono"},formatter:v=>fmt(v),offset:2}},
    {type:"line",label:"Avg",data:data.map(m=>m.total),
      borderColor:th.income,backgroundColor:"transparent",borderWidth:2,
      pointRadius:isMobile?2:3,pointBackgroundColor:th.income,tension:0.35,yAxisID:"y",datalabels:{display:false}}
  ]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:"index"},layout:{padding:{top:isMobile?22:28}},
    plugins:{legend:{display:false},tooltip:{backgroundColor:th.tooltip,borderColor:th.border,borderWidth:1,titleColor:th.text,bodyColor:th.muted,padding:10,callbacks:{label:ctx=>"  "+ctx.dataset.label+": "+fmtFull(ctx.parsed.y)}}},
    scales:{x:{ticks:{color:th.muted,font:{size:isMobile?8:10},maxRotation:35},grid:{color:th.grid}},
      y:{ticks:{color:th.muted,font:{size:isMobile?8:10},callback:v=>fmt(v)},grid:{color:th.grid}}}}};
}

// ── categoryConfig: barras com espaçamento adequado ───────────
function categoryConfig(cats, isZoom, isMobile){
  const th = TH();
  const sorted = [...cats].reverse();
  const pal = P();

  // User-friendly sizing: smaller, cleaner labels in zoom; no heavy mono font.
  const labelFont = isZoom ? 11 : isMobile ? 10 : 11;
  const valueFont = isZoom ? 11 : isMobile ? 9 : 10;
  const barMax = isZoom ? 28 : isMobile ? 22 : 24;
  const maxLabel = isZoom ? 24 : isMobile ? 14 : 20;

  return {
    type: "bar",
    data: {
      labels: sorted.map(c => c.cat),
      datasets: [{
        label: "Amount",
        data: sorted.map(c => c.total),
        backgroundColor: sorted.map((_, i) => pal[i % pal.length] + "d9"),
        borderColor: sorted.map((_, i) => pal[i % pal.length]),
        borderWidth: 0,
        borderRadius: 8,
        barPercentage: isZoom ? 0.72 : isMobile ? 0.68 : 0.70,
        categoryPercentage: isZoom ? 0.82 : isMobile ? 0.78 : 0.80,
        maxBarThickness: barMax,
        datalabels: {
          anchor: "end",
          align: "right",
          clamp: true,
          clip: false,
          color: th.text,
          font: { size: valueFont, weight: "700", family: "Inter, system-ui, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif" },
          formatter: v => fmtFull(v),
          offset: isZoom ? 10 : 8
        }
      }]
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      devicePixelRatio: Math.min(window.devicePixelRatio || 1, 2),
      layout: {
        padding: {
          left: isZoom ? 16 : 4,
          right: isZoom ? 115 : isMobile ? 88 : 105,
          top: isZoom ? 10 : 6,
          bottom: isZoom ? 12 : 6
        }
      },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: th.tooltip,
          borderColor: th.border,
          borderWidth: 1,
          titleColor: th.text,
          bodyColor: th.muted,
          padding: 10,
          callbacks: { label: ctx => "  " + fmtFull(ctx.parsed.x) }
        }
      },
      scales: {
        x: {
          beginAtZero: true,
          grace: isZoom ? "8%" : "10%",
          ticks: {
            color: th.muted,
            font: { size: isZoom ? 10 : isMobile ? 9 : 10, family: "Inter, system-ui, sans-serif" },
            callback: v => fmt(v),
            maxTicksLimit: isZoom ? 5 : isMobile ? 4 : 5
          },
          grid: { color: th.grid, drawBorder: false }
        },
        y: {
          ticks: {
            color: th.text,
            font: { size: labelFont, weight: "650", family: "Inter, system-ui, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif" },
            padding: isZoom ? 8 : 9,
            callback: function(value){
              const label = this.getLabelForValue(value);
              return label.length > maxLabel ? label.substring(0, maxLabel) + "…" : label;
            }
          },
          grid: { display: false, drawBorder: false }
        }
      }
    }
  };
}

function trendConfig(all, isZoom, isMobile){
  const th = TH();
  const avg = all.reduce((s, m) => s + m.total, 0) / (all.length || 1);
  const ptColors = all.map(m => m.total > avg ? th.expense : th.income);

  return {
    type: "line",
    data: {
      labels: all.map(m => m.label),
      datasets: [
        {
          label: "Monthly",
          data: all.map(m => m.total),
          borderColor: th.trend,
          backgroundColor: th.trend + "20",
          borderWidth: isMobile ? 2 : 3,
          pointRadius: all.map((_, i) => i === all.length - 1 ? 6 : isMobile ? 3 : 4),
          pointBackgroundColor: ptColors,
          pointBorderColor: ptColors,
          pointHoverRadius: 8,
          tension: 0.42,
          fill: true,
          datalabels: {
            display: function(ctx){
              const vals = ctx.dataset.data;
              const max = Math.max(...vals);
              const min = Math.min(...vals);
              const v = vals[ctx.dataIndex];
              return v === max || v === min || ctx.dataIndex === vals.length - 1;
            },
            anchor: "end",
            align: "top",
            color: th.text,
            backgroundColor: th.surface,
            borderColor: th.border,
            borderWidth: 1,
            borderRadius: 6,
            padding: 4,
            font: { size: isZoom ? 11 : isMobile ? 9 : 10, weight: "800", family: "JetBrains Mono" },
            formatter: v => fmt(v),
            offset: 6,
            clamp: true
          }
        },
        {
          label: "Avg " + fmt(Math.round(avg)),
          data: all.map(() => Math.round(avg)),
          borderColor: th.avg + "cc",
          borderWidth: 2,
          borderDash: [6,5],
          pointRadius: 0,
          fill: false,
          datalabels: { display: false }
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index" },
      layout: { padding: { top: isZoom ? 34 : 28, right: 18, left: 8, bottom: 8 } },
      plugins: {
        legend: {
          position: "top",
          align: "center",
          labels: { color: th.muted, font: { size: isMobile ? 10 : 12, weight: "600" }, boxWidth: 10, usePointStyle: true, padding: 18 }
        },
        tooltip: {
          backgroundColor: th.tooltip,
          borderColor: th.border,
          borderWidth: 1,
          titleColor: th.text,
          bodyColor: th.muted,
          padding: 10,
          callbacks: { label: ctx => "  " + ctx.dataset.label + ": " + fmtFull(ctx.parsed.y) }
        }
      },
      scales: {
        x: {
          ticks: { color: th.muted, font: { size: isZoom ? 11 : isMobile ? 9 : 10 }, maxTicksLimit: isMobile ? 5 : 8, maxRotation: isMobile ? 45 : 35, autoSkip: true },
          grid: { color: th.grid }
        },
        y: {
          beginAtZero: true,
          grace: "15%",
          ticks: { color: th.muted, font: { size: isZoom ? 11 : isMobile ? 9 : 10 }, callback: v => fmt(v), maxTicksLimit: 5 },
          grid: { color: th.grid }
        }
      }
    }
  };
}

function buildHeatmapHTML(hd,large){
  if(!hd||!hd.cats.length)return'<p style="color:var(--muted);padding:16px">No data</p>';
  const th=TH();
  const stats=hd.cats.map((_,ci)=>{const vals=hd.values[ci].filter(v=>v>0);return{max:Math.max(...vals,1),min:Math.min(...vals,Infinity),avg:vals.reduce((s,v)=>s+v,0)/(vals.length||1)};});
  const fs=large?"11px":"9px",pd=large?"7px 10px":"4px 6px";
  let h='<div style="display:flex;align-items:center;gap:8px;padding:6px 8px 4px;font-size:9px;color:'+th.muted+'">';
  h+='<span>Intensity per category:</span>';
  const gradStops=["rgba(99,102,241,.12)","rgba(99,102,241,.4)","rgba(99,102,241,.7)","rgba(99,102,241,1)"];
  h+='<div style="display:flex;gap:2px;align-items:center">';
  gradStops.forEach(c=>{h+='<div style="width:14px;height:10px;border-radius:2px;background:'+c+'"></div>';});
  h+='<span style="margin-left:3px">Low → High</span></div><span style="margin-left:auto;font-size:8px">🟢 below avg · 🔴 above avg</span></div>';
  h+='<table class="hm-table"><thead><tr><th style="min-width:'+(large?"120px":"80px")+'">Category</th>';
  hd.months.forEach(m=>{h+='<th style="font-size:'+fs+'">'+m+'</th>';});
  h+='<th style="font-size:'+(large?"9px":"8px")+'">Avg/mo</th></tr></thead><tbody>';
  hd.cats.forEach((cat,ci)=>{
    const st=stats[ci],icon=CAT_ICONS[cat]||"";
    h+='<tr><td class="hm-cat" style="font-size:'+(large?"11px":"9px")+'">'+icon+' '+cat+'</td>';
    hd.values[ci].forEach((v,mi)=>{
      if(!v){h+='<td style="background:var(--surf2);color:var(--muted);padding:'+pd+';border-radius:4px;font-size:'+fs+'">—</td>';return;}
      const intensity=Math.max(0.08,v/st.max);
      const isAboveAvg=v>st.avg*1.1,isBelowAvg=v<st.avg*0.9&&v>0,isLatest=mi===hd.values[ci].length-1;
      const fg=intensity>0.5?"#fff":(currentTheme==="light"?"#1e1b4b":"#c7d2fe");
      let ring=isLatest?"outline:2px solid "+th.income+";outline-offset:-2px;":"";
      let badge=isAboveAvg?"🔴":isBelowAvg?"🟢":"";
      h+='<td style="background:rgba(99,102,241,'+intensity.toFixed(2)+');color:'+fg+';padding:'+pd+';font-weight:600;border-radius:4px;font-size:'+fs+';'+ring+'" title="'+cat+': '+fmtFull(v)+' (avg: '+fmtFull(Math.round(st.avg))+')">'+(large?fmtFull(v):fmt(v));
      if(badge)h+='<span style="font-size:7px;margin-left:2px">'+badge+'</span>';
      h+='</td>';
    });
    h+='<td style="background:'+th.grid+';color:'+th.muted+';padding:'+pd+';font-size:'+(large?"10px":"8px")+';;border-radius:4px;font-weight:600;text-align:center">'+fmt(Math.round(st.avg))+'</td></tr>';
  });
  return h+'</tbody></table>';
}

// ── Desktop chart updaters ────────────────────────────────────
function updateMonthlyChart(){
  if(!allData)return;
  const n=activePills,data=n>=999?allData.allMonthlyTotals:allData.allMonthlyTotals.slice(-n);
  if(charts.monthly){charts.monthly.destroy();delete charts.monthly;}
  charts.monthly=new Chart(document.getElementById("chart-monthly").getContext("2d"),monthlyConfig(data,false));
}

function updateCategoryChart(){
  if(!allData) return;
  const key = document.getElementById("sel-cat-month").value;
  const cats = key === "all" ? allData.categoryTotals.slice(0,10) : (allData.monthCatData[key] || []).slice(0,10);
  const wrap = document.getElementById("chart-category-wrap");
  const inner = document.getElementById("chart-category-inner");
  if(wrap && inner){
    // Menos espaçado: mantém altura fixa confortável e scroll interno.
    wrap.style.height = "420px";
    wrap.style.maxHeight = "420px";
    wrap.style.overflowY = "auto";
    inner.style.height = Math.max(420, cats.length * 34 + 60) + "px";
  }
  if(charts.category){charts.category.destroy(); delete charts.category;}
  charts.category = new Chart(document.getElementById("chart-category").getContext("2d"), categoryConfig(cats,false,false));
}

function renderTrend(){
  if(charts.trend){charts.trend.destroy();delete charts.trend;}
  charts.trend=new Chart(document.getElementById("chart-trend").getContext("2d"),trendConfig(allData.allMonthlyTotals,false,false));
}
function renderHeatmap(){document.getElementById("hm-main").innerHTML=buildHeatmapHTML(allData.heatmapData,false);}

// ── Mobile chart updaters ─────────────────────────────────────
function updateMonthlyChartMobile(){
  if(!allData)return;
  const n=activePills,data=n>=999?allData.allMonthlyTotals:allData.allMonthlyTotals.slice(-n);
  if(charts.monthlyM){charts.monthlyM.destroy();delete charts.monthlyM;}
  const el=document.getElementById("chart-monthly-m");
  if(el)charts.monthlyM=new Chart(el.getContext("2d"),monthlyConfig(data,true));
}

function updateCategoryChartMobile(){
  if(!allData) return;
  const sel = document.getElementById("sel-cat-month-m");
  const key = sel ? sel.value : "all";
  const cats = key === "all" ? allData.categoryTotals.slice(0,10) : (allData.monthCatData[key] || []).slice(0,10);
  const wrap = document.getElementById("chart-category-wrap-m");
  const inner = document.getElementById("chart-category-inner-m");
  if(wrap && inner){
    const neededHeight = Math.max(340, cats.length * 32 + 60);
    wrap.style.height = Math.min(neededHeight, 430) + "px";
    wrap.style.maxHeight = "430px";
    wrap.style.overflowY = neededHeight > 430 ? "auto" : "hidden";
    inner.style.height = neededHeight + "px";
  }
  if(charts.categoryM){charts.categoryM.destroy(); delete charts.categoryM;}
  const el = document.getElementById("chart-category-m");
  if(el) charts.categoryM = new Chart(el.getContext("2d"), categoryConfig(cats,false,true));
}

function renderTrendMobile(){
  if(charts.trendM){charts.trendM.destroy();delete charts.trendM;}
  const el=document.getElementById("chart-trend-m");
  if(el)charts.trendM=new Chart(el.getContext("2d"),trendConfig(allData.allMonthlyTotals,false,true));
}
function renderHeatmapMobile(){const el=document.getElementById("hm-mobile");if(el)el.innerHTML=buildHeatmapHTML(allData.heatmapData,false);}

function populateMonthSelectors(){
  [document.getElementById("sel-cat-month"),document.getElementById("sel-cat-month-m")].forEach(sel=>{
    if(!sel)return;
    while(sel.options.length>0)sel.remove(0);
    var optAll=document.createElement("option");optAll.value="all";optAll.textContent="All time";sel.appendChild(optAll);
    allData.availableMonths.slice().reverse().forEach(function(m){var o=document.createElement("option");o.value=m.key;o.textContent=m.label;sel.appendChild(o);});
  });
}

function renderAllCharts(){
  populateMonthSelectors();
  if(IS_MOBILE){
    updateMonthlyChartMobile();updateCategoryChartMobile();renderTrendMobile();renderHeatmapMobile();
  }else{
    updateMonthlyChart();updateCategoryChart();renderTrend();renderHeatmap();
  }
}

let _resizeTimer=null;
window.addEventListener("resize", function(){
  clearTimeout(_resizeTimer);
  _resizeTimer=setTimeout(function(){
    const wasMobile=IS_MOBILE;
    IS_MOBILE = window.innerWidth <= 700 || (window.screen && window.screen.width <= 600);
    document.body.classList.toggle("is-mobile", IS_MOBILE);
    if(allData){destroyAll();renderAllCharts();}
  }, 180);
});

// ── Budget ────────────────────────────────────────────────────
// ── Spending Layers (Fixed / Variable / Lifestyle) ──────────────
var LAYER_META={
  Fixed:    {icon:"🔒",label:"Fixed",    tag:"LOCKED",   color:"var(--muted)"},
  Variable: {icon:"🛒",label:"Variable", tag:"NEEDS",    color:"var(--orange)"},
  Lifestyle:{icon:"🎉",label:"Lifestyle",tag:"LIFESTYLE",color:"var(--purple)"}
};
var activeLayer=null;
var catLayerMap={};

function renderLayersPanel(layerData){
  const wrap=document.getElementById("layers-wrap");
  if(!layerData){wrap.style.display="none";activeLayer=null;return;}
  wrap.style.display="block";

  catLayerMap={};
  ["Fixed","Variable","Lifestyle"].forEach(function(k){
    (layerData[k].cats||[]).forEach(function(c){catLayerMap[c.cat]=k;});
  });

  document.getElementById("lh-controllable").textContent=fmtFull(layerData.controllable||0);
  const total=layerData.totalSpent||0;

  document.getElementById("layers-cards").innerHTML=["Fixed","Variable","Lifestyle"].map(function(k){
    const L=layerData[k],m=LAYER_META[k];
    const pctTotal=total>0?Math.round((L.spent/total)*100):0;
    const cls="layer-card layer-"+k.toLowerCase()+(activeLayer===k?" active":"")+(k==="Fixed"?" locked":"");
    const sub=L.budget>0?("of "+fmtFull(L.budget)+" budget"):(L.cats.length+" categories");
    return '<div class="'+cls+'" data-layer="'+k+'">'+
      '<div class="lc-top"><span class="lc-icon">'+m.icon+'</span><span class="lc-name">'+m.label+'</span><span class="lc-tag">'+m.tag+'</span></div>'+
      '<div class="lc-amount">'+fmtFull(L.spent)+'</div>'+
      '<div class="lc-bar-bg"><div class="lc-bar-fill" style="width:'+pctTotal+'%;background:'+m.color+'"></div></div>'+
      '<div class="lc-sub"><span>'+pctTotal+'% of spend</span><span>'+sub+'</span></div>'+
    '</div>';
  }).join("");
}

function selectLayer(layer){
  activeLayer=(activeLayer===layer)?null:layer;
  if(allData){
    renderLayersPanel(allData.layerData);
    renderBudgetPanel(allData.budgetData);
  }
  var fl=document.getElementById("budget-filter-lbl");
  if(fl)fl.textContent=activeLayer?("  ·  filtered: "+activeLayer):"";
}

function renderBudgetPanel(budgetData){
  const grid=document.getElementById("budget-grid");
  if(!budgetData||budgetData.length===0){grid.innerHTML='<div style="grid-column:1/-1;padding:40px;text-align:center;color:var(--muted)">No budgets set.<br><br>Go to menu <strong>💰 Budget Importer → 💵 Setup Budget Sheet</strong>.</div>';return;}
  const mk=allData?allData.currentMonthKey:"";
  let rows=budgetData;
  if(activeLayer)rows=budgetData.filter(function(item){return catLayerMap[item.cat]===activeLayer;});
  if(rows.length===0){grid.innerHTML='<div style="grid-column:1/-1;padding:30px;text-align:center;color:var(--muted)">No budgeted categories in <strong>'+activeLayer+'</strong>.<br><br>Set a budget for these categories, or click the layer card again to show all.</div>';return;}
  grid.innerHTML=rows.map(function(item){
    const pct=Math.min(item.pct,100);
    const barColor=item.status==="over"?"var(--red)":item.status==="warning"?"var(--yellow)":"var(--green)";
    const stClass=item.status==="over"?"st-over":item.status==="warning"?"st-warning":item.budget>0?"st-ok":"st-none";
    const stLabel=item.status==="over"?"OVER":item.status==="warning"?"⚠ "+item.pct+"%":item.budget>0?item.pct+"%":"NO LIMIT";
    var suggBadge="";
    if(item.suggestion){
      var adjTxt="Adjust to "+fmtFull(item.suggestion);
      var spentTxt="💡 Spent "+fmtFull(item.spent)+" · Projected: "+fmtFull(item.suggestion);
      suggBadge='<div class="suggestion-bar"><span class="sugg-text">'+spentTxt+'</span><button class="sugg-btn" data-sugg-cat="'+item.cat+'" data-sugg-val="'+item.suggestion+'">'+adjTxt+'</button></div>';
    }
    return '<div class="bcard" data-drill-cat="'+item.cat+'" data-drill-mk="'+mk+'" data-drill-w="0">'+
      '<div class="bcard-top"><div class="bcat">'+(CAT_ICONS[item.cat]||"💳")+" "+item.cat+'</div><div class="bstatus '+stClass+'">'+stLabel+'</div></div>'+
      '<div class="bbar-bg"><div class="bbar-fill" style="width:'+pct+'%;background:'+barColor+'"></div></div>'+
      '<div class="bcard-nums"><span class="spent">'+fmtFull(item.spent)+'</span><span>of '+(item.budget>0?fmtFull(item.budget):"no limit")+'</span></div>'+
      suggBadge+
      '<div class="bcard-edit" onclick="event.stopPropagation()"><label>Budget $</label><input class="budget-input" type="number" inputmode="numeric" data-cat="'+item.cat+'" value="'+(item.budget||"")+'" placeholder="0"/></div>'+
      '<div class="bcard-hint">👆 Click to view transactions</div>'+
    '</div>';
  }).join("");
}

// ── Income / Remaining ────────────────────────────────────────
function syncIncome(el){
  var d=document.getElementById("income-input");
  if(d)d.value=el.value;
  _doUpdateRemaining();
}
function updateRemaining(){_doUpdateRemaining();}
function _doUpdateRemaining(){
  var income=parseFloat(document.getElementById("income-input").value||"0")||parseFloat((document.getElementById("income-input-mobile")||{}).value||"0")||0;
  if(!allData)return;
  const all=allData.allMonthlyTotals,curr=all.length?all[all.length-1].total:0;
  const remaining=income-curr,pct=income>0?Math.min(100,(curr/income)*100):0;
  document.getElementById("kpi-remaining").textContent=income>0?fmtFull(remaining):"—";
  const card=document.getElementById("remaining-card");
  if(income>0){
    card.style.display="block";
    document.getElementById("rem-amount").textContent=fmtFull(remaining);
    document.getElementById("rem-amount").style.color=remaining>=0?"var(--green)":"var(--red)";
    document.getElementById("rem-bar").style.width=pct+"%";
    document.getElementById("rem-bar").style.background=pct<70?"var(--green)":pct<90?"var(--yellow)":"var(--red)";
    document.getElementById("rem-spent-lbl").textContent="Spent: "+fmtFull(curr);
    document.getElementById("rem-income-lbl").textContent="Income: "+fmtFull(income);
    const now=new Date(),daysInMonth=new Date(now.getFullYear(),now.getMonth()+1,0).getDate(),daysPassed=now.getDate();
    document.getElementById("rem-projection").textContent="📍 At this pace, projected spend this month: "+fmtFull(Math.round((curr/daysPassed)*daysInMonth*100)/100);
  }else{card.style.display="none";}
}

function applySuggestion(btn){
  var cat=btn.getAttribute("data-sugg-cat");var newBudget=parseFloat(btn.getAttribute("data-sugg-val"));
  if(!cat||!newBudget)return;
  var targetInput=null;
  document.querySelectorAll(".budget-input").forEach(function(inp){if(inp.getAttribute("data-cat")===cat)targetInput=inp;});
  if(!targetInput)return;
  targetInput.value=newBudget;btn.textContent="⏳";btn.disabled=true;
  google.script.run
    .withSuccessHandler(function(res){
      if(res.ok){var bar=btn.closest(".suggestion-bar");if(bar)bar.innerHTML='<span style="color:var(--green);font-size:10px">✅ Budget updated to '+fmtFull(newBudget)+'</span>';targetInput.value=newBudget;}
      else{btn.textContent="Adjust to "+fmtFull(newBudget);btn.disabled=false;}
    })
    .withFailureHandler(function(){btn.textContent="Adjust to "+fmtFull(newBudget);btn.disabled=false;})
    .saveBudgetsFromDashboard([{cat:cat,budget:newBudget}]);
}

function saveBudgets(){
  var inputs=document.querySelectorAll(".budget-input");
  if(!inputs||inputs.length===0){alert("No budget fields found.");return;}
  var budgetArray=[];
  inputs.forEach(function(inp){var cat=inp.getAttribute("data-cat");var val=parseFloat(inp.value)||0;if(cat)budgetArray.push({cat:cat,budget:val});});
  var btn=document.getElementById("save-budget-btn");
  btn.textContent="⏳ Saving...";btn.disabled=true;
  google.script.run
    .withSuccessHandler(function(res){btn.disabled=false;btn.textContent=res.ok?"✅ Saved!":"❌ "+(res.error||"Error");setTimeout(function(){btn.textContent="💾 Save Budgets";},2500);if(res.ok)loadData();})
    .withFailureHandler(function(err){btn.disabled=false;btn.textContent="❌ "+(err&&err.message||"Error");setTimeout(function(){btn.textContent="💾 Save Budgets";},2500);})
    .saveBudgetsFromDashboard(budgetArray);
}

// ── Leaks ─────────────────────────────────────────────────────
function renderLeaksPanel(leaks){
  const grid=document.getElementById("leak-grid");
  if(!leaks||leaks.length===0){grid.innerHTML='<div style="padding:40px;color:var(--muted);grid-column:1/-1;text-align:center">✅ No spending leaks detected.</div>';return;}
  const lvl={"high":{bg:"rgba(248,113,113,.12)",bd:"rgba(248,113,113,.3)",badge:"var(--red)",label:"HIGH"},"medium":{bg:"rgba(251,191,36,.10)",bd:"rgba(251,191,36,.3)",badge:"var(--yellow)",label:"MEDIUM"},"low":{bg:"rgba(99,102,241,.08)",bd:"rgba(99,102,241,.25)",badge:"var(--accent)",label:"LOW"}};
  const w=parseInt(document.getElementById("leak-window").value)||3;
  grid.innerHTML=leaks.map(function(leak){
    const icon=CAT_ICONS[leak.cat]||"💸";const lc=lvl[leak.level]||lvl.low;
    const freqTxt=leak.freqPerWeek>=1?leak.freqPerWeek.toFixed(1)+"x/week":Math.round(leak.freqPerWeek*30)+"x/month";
    const yearly=Math.round(leak.avgPerMonth*12);
    return '<div class="leak-card" data-drill-cat="'+leak.cat+'" data-drill-mk="all" data-drill-w="'+w+'" style="background:'+lc.bg+';border:1px solid '+lc.bd+'">'+
      '<div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;"><span style="font-size:24px">'+icon+'</span>'+
      '<div style="flex:1"><div style="font-size:13px;font-weight:700">'+leak.cat+'</div><div style="font-size:10px;color:var(--muted)">'+freqTxt+' · '+leak.txCount+' transactions · '+leak.monthCount+' months</div></div>'+
      '<div style="text-align:right"><div style="font-size:9px;font-weight:700;padding:2px 8px;border-radius:20px;background:var(--surf2);color:'+lc.badge+'">'+lc.label+'</div><div style="font-size:9px;color:var(--muted);margin-top:4px">click to inspect</div></div></div>'+
      '<div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:5px;text-align:center;">'+
        '<div style="background:var(--surf2);border-radius:8px;padding:7px 4px;"><div style="font-size:8px;color:var(--muted);margin-bottom:2px;text-transform:uppercase">Count</div><div style="font-size:15px;font-weight:700;color:var(--accent)">'+leak.txCount+'</div></div>'+
        '<div style="background:var(--surf2);border-radius:8px;padding:7px 4px;"><div style="font-size:8px;color:var(--muted);margin-bottom:2px;text-transform:uppercase">Avg/charge</div><div style="font-size:13px;font-weight:700;font-family:monospace;color:var(--yellow)">'+fmtFull(leak.avgPerTx)+'</div></div>'+
        '<div style="background:var(--surf2);border-radius:8px;padding:7px 4px;"><div style="font-size:8px;color:var(--muted);margin-bottom:2px;text-transform:uppercase">Period total</div><div style="font-size:13px;font-weight:700;font-family:monospace;color:var(--red)">'+fmtFull(leak.totalAmt)+'</div></div>'+
        '<div style="background:var(--surf2);border-radius:8px;padding:7px 4px;"><div style="font-size:8px;color:var(--muted);margin-bottom:2px;text-transform:uppercase">If yearly</div><div style="font-size:13px;font-weight:700;font-family:monospace;color:var(--orange)">'+fmt(yearly)+'</div></div>'+
      '</div></div>';
  }).join("");
}

function reloadLeaks(){
  if(!allData)return;
  const w=parseInt(document.getElementById("leak-window").value)||3;
  document.getElementById("leak-grid").innerHTML='<div style="padding:20px;color:var(--muted)">⏳ Analyzing...</div>';
  google.script.run
    .withSuccessHandler(function(res){if(res.ok&&res.data)renderLeaksPanel(res.data.leaks);})
    .withFailureHandler(function(){document.getElementById("leak-grid").innerHTML='<div style="padding:20px;color:var(--red)">❌ Error.</div>';})
    .getDashboardData(w, false);
}

// ── Drill-down ────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded",function(){
  initMobileClass();
  document.addEventListener("click",function(e){
    var sugg=e.target.closest(".sugg-btn");
    if(sugg){e.stopPropagation();applySuggestion(sugg);return;}
    if(e.target.closest(".suggestion-bar")){e.stopPropagation();return;}
    var lyr=e.target.closest("[data-layer]");
    if(lyr){selectLayer(lyr.getAttribute("data-layer"));return;}
    var el=e.target.closest("[data-drill-cat]");
    if(el&&!e.target.closest(".bcard-edit")&&!e.target.closest(".budget-input")){
      openDrill(el.getAttribute("data-drill-cat"),el.getAttribute("data-drill-mk")||"",parseInt(el.getAttribute("data-drill-w"))||0);
    }
    if(e.target===document.getElementById("drill-overlay"))closeDrill();
  });
  if(IS_MOBILE){
    window.addEventListener("popstate",function(){
      if(document.getElementById("drill-overlay").style.display!=="none"){closeDrill();history.pushState(null,"",location.href);}
    });
  }
});

function openDrill(cat,monthKey,windowMonths){
  var icon=CAT_ICONS[cat]||"📋";
  document.getElementById("drill-icon").textContent=icon;
  document.getElementById("drill-name").textContent=cat;
  var sub=monthKey==="all"?"Last "+windowMonths+" months":(allData&&allData.availableMonths.find(function(m){return m.key===monthKey;})||{label:"Current month"}).label;
  document.getElementById("drill-sub").textContent=sub;
  document.getElementById("ds-count").textContent="—";document.getElementById("ds-total").textContent="—";
  document.getElementById("ds-avg").textContent="—";document.getElementById("ds-max").textContent="—";
  document.getElementById("drill-table").style.display="none";
  document.getElementById("drill-overlay").style.display="flex";
  if(IS_MOBILE)history.pushState(null,"",location.href);

  var cachedRows=(allData&&allData._txByCat&&allData._txByCat[cat])||[];
  if(cachedRows.length>0){
    document.getElementById("drill-loading").style.display="none";
    _renderDrillTable(cat,monthKey,windowMonths);
    return;
  }

  document.getElementById("drill-loading").textContent="⏳ Loading transactions... first time only, then it is cached.";
  document.getElementById("drill-loading").style.display="block";
  google.script.run
    .withSuccessHandler(function(res){
      if(!res||!res.ok){
        document.getElementById("drill-loading").textContent="❌ "+(res&&res.error||"Could not load transactions.");
        return;
      }
      allData._txByCat=allData._txByCat||{};
      allData._txByCat[cat]=res.txs||[];
      _renderDrillTable(cat,monthKey,windowMonths);
    })
    .withFailureHandler(function(e){
      document.getElementById("drill-loading").textContent="❌ "+(e&&e.message||"Failed.");
    })
    .getTxForCategory(cat,monthKey||"all",windowMonths||0);
}

function _renderDrillTable(cat,monthKey,windowMonths){
  var allRows=(allData&&allData._txByCat&&allData._txByCat[cat])||[];
  // Build indexed pairs so the original position in allRows is preserved after filtering.
  // Bug fix: previously idx came from the filtered array, but reclassifyRow spliced the full array,
  // causing the wrong transaction to be removed from the local cache.
  var indexed=allRows.map(function(r,i){return{row:r,origIdx:i};});
  var filtered;
  if(monthKey==="all"&&windowMonths){
    var now=new Date();var limit=new Date(now.getFullYear(),now.getMonth()-(windowMonths-1),1);
    var limitMk=limit.getFullYear()+"-"+String(limit.getMonth()+1).padStart(2,"0");
    filtered=indexed.filter(function(e){return e.row.monthKey>=limitMk;});
  }else if(monthKey){filtered=indexed.filter(function(e){return e.row.monthKey===monthKey;});}
  else{filtered=indexed;}
  if(!filtered||filtered.length===0){document.getElementById("drill-loading").textContent="No transactions found for this period.";document.getElementById("drill-loading").style.display="block";return;}
  document.getElementById("drill-loading").style.display="none";
  var total=filtered.reduce(function(s,e){return s+e.row.amount;},0);
  var avg=total/filtered.length,max=Math.max.apply(null,filtered.map(function(e){return e.row.amount;}));
  document.getElementById("ds-count").textContent=filtered.length;
  document.getElementById("ds-total").textContent=fmtFull(total);
  document.getElementById("ds-avg").textContent=fmtFull(avg);
  document.getElementById("ds-max").textContent=fmtFull(max);
  var catOpts=ALL_CATS.map(function(c){return '<option value="'+c+'"'+(c===cat?' selected':'')+'>'+c+'</option>';}).join('');
  document.getElementById("drill-tbody").innerHTML=filtered.map(function(e){
    var r=e.row,origIdx=e.origIdx;
    return "<tr><td class='td-date'>"+esc(r.date)+"</td><td class='td-desc' title='"+esc(r.desc)+"'>"+esc(r.desc)+"</td><td class='td-amt'>"+fmtFull(r.amount)+"</td><td class='td-bank'>"+esc(r.bank)+"</td><td class='td-recat'><select class='recat-sel' data-idx='"+origIdx+"' data-cat='"+esc(cat)+"' data-mk='"+esc(r.monthKey)+"' onchange='reclassifyRow(this,"+JSON.stringify({desc:r.desc,amount:r.amount,monthKey:r.monthKey,bank:r.bank})+")'>"+catOpts+"</select></td></tr>";
  }).join("");
  document.getElementById("drill-table").style.display="table";
}

function reclassifyRow(sel,txInfo){
  var newCat=sel.value;txInfo.newCategory=newCat;sel.classList.add("saving");sel.disabled=true;
  google.script.run
    .withSuccessHandler(function(res){
      sel.classList.remove("saving");sel.disabled=false;
      if(res.ok){sel.classList.add("saved");var idx=parseInt(sel.getAttribute("data-idx"));var oldCat=sel.getAttribute("data-cat");
        if(allData&&allData._txByCat){var rows=allData._txByCat[oldCat]||[];var tx=rows[idx];if(tx){rows.splice(idx,1);if(!allData._txByCat[newCat])allData._txByCat[newCat]=[];tx.cat=newCat;allData._txByCat[newCat].unshift(tx);}}
        setTimeout(function(){sel.classList.remove("saved");},1500);
      }else{sel.value=sel.getAttribute("data-cat");alert("Error: "+(res.error||"Could not save."));}
    })
    .withFailureHandler(function(e){sel.classList.remove("saving");sel.disabled=false;sel.value=sel.getAttribute("data-cat");alert("Error: "+(e&&e.message||"Failed."));})
    .reclassifyTransaction(txInfo);
}
function closeDrill(){document.getElementById("drill-overlay").style.display="none";}

// ── Zoom ──────────────────────────────────────────────────────

function getActiveCategoryKey(){
  const desktopSel=document.getElementById("sel-cat-month");
  const mobileSel=document.getElementById("sel-cat-month-m");
  return (IS_MOBILE && mobileSel ? mobileSel.value : desktopSel ? desktopSel.value : "all") || "all";
}
function getCategoryRowsForKey(key){
  key = key || "all";
  return key === "all" ? allData.categoryTotals.slice(0,10) : (allData.monthCatData[key] || []).slice(0,10);
}
function syncCategorySelectors(key){
  [document.getElementById("sel-cat-month"),document.getElementById("sel-cat-month-m")].forEach(function(sel){
    if(sel) sel.value = key || "all";
  });
}
function renderZoomCategory(key){
  const cw = document.getElementById("zoom-cw");
  const canvas = document.getElementById("zoom-canvas");
  const cats = getCategoryRowsForKey(key);

  const chartHeight = Math.max(420, cats.length * 58 + 120);

  cw.style.height = "100%";
  cw.style.overflowY = "auto";
  cw.style.overflowX = "hidden";

  canvas.style.width = "100%";
  canvas.style.height = chartHeight + "px";
  canvas.style.minHeight = chartHeight + "px";

  if(zoomChart){
    zoomChart.destroy();
    zoomChart = null;
  }

  zoomChart = new Chart(
    canvas.getContext("2d"),
    categoryConfig(cats, true, false)
  );
}
function buildZoomCategoryFilter(){
  const holder=document.getElementById("zoom-controls");
  if(!holder) return;
  holder.innerHTML="";
  const sel=document.createElement("select");
  sel.id="zoom-cat-month";
  sel.className="fsel";
  sel.innerHTML=document.getElementById("sel-cat-month").innerHTML;
  sel.value=getActiveCategoryKey();
  sel.onchange=function(){
    syncCategorySelectors(sel.value);
    updateCategoryChart();
    if(IS_MOBILE) updateCategoryChartMobile();
    renderZoomCategory(sel.value);
  };
  holder.appendChild(sel);
}
function clearZoomControls(){
  const holder=document.getElementById("zoom-controls");
  if(holder) holder.innerHTML="";
}

function openZoom(type){
  const titles={
    monthly:"📅 Monthly Expenses",
    category:"🏷️ Top 10 Categories",
    trend:"📈 Spending Trend",
    heatmap:"🔥 Category Heatmap"
  };

  document.getElementById("zoom-title").textContent=titles[type]||type;
  clearZoomControls();

  if(zoomChart){
    zoomChart.destroy();
    zoomChart=null;
  }

  const cw=document.getElementById("zoom-cw");
  const hm=document.getElementById("zoom-hm");
  const canvas=document.getElementById("zoom-canvas");

  cw.style.display="none";
  hm.style.display="none";
  cw.style.height="100%";
  cw.style.overflow="auto";
  canvas.style.setProperty("--zoom-canvas-height","100%");

  if(type==="heatmap"){
    hm.style.display="block";
    hm.innerHTML=buildHeatmapHTML(allData.heatmapData,true);
  } else {
    cw.style.display="block";

    let cfg;

    if(type==="monthly"){
      const n=activePills;
      const data=n>=999 ? allData.allMonthlyTotals : allData.allMonthlyTotals.slice(-n);
      cfg=monthlyConfig(data,false);
    }

    else if(type==="category"){
      buildZoomCategoryFilter();
      cfg = null;

      document.getElementById("zoom-overlay").classList.add("open");

      setTimeout(function(){
        renderZoomCategory(getActiveCategoryKey());
      }, 80);

      return;
    }

    else if(type==="trend"){
      cfg=trendConfig(allData.allMonthlyTotals,true,false);
    }

    if(cfg){
      zoomChart=new Chart(canvas.getContext("2d"),cfg);
    }
  }

  document.getElementById("zoom-overlay").classList.add("open");
}
function closeZoom(){document.getElementById("zoom-overlay").classList.remove("open");if(zoomChart){zoomChart.destroy();zoomChart=null;}}
function closeZoomOutside(e){if(e.target===document.getElementById("zoom-overlay"))closeZoom();}


function safeCloseDashboard(){
  try {
    if (typeof google !== "undefined" && google.script && google.script.host) {
      google.script.host.close();
      return;
    }
  } catch(e) {}
  try { window.close(); } catch(e) {}
}


function warmTxCacheForCurrentViews(){
  if(!allData) return;
  allData._txByCat = allData._txByCat || {};
  // Collect categories to pre-warm, capped to avoid too many concurrent server calls.
  // Budget categories are highest priority (visible immediately); leaks second.
  const cats = [];
  (allData.budgetData || []).slice(0, 8).forEach(function(x){ if(x && x.cat && !cats.includes(x.cat)) cats.push(x.cat); });
  (allData.leaks || []).slice(0, 6).forEach(function(x){ if(x && x.cat && !cats.includes(x.cat)) cats.push(x.cat); });
  const toFetch = cats.filter(function(cat){ return cat && !(allData._txByCat[cat] && allData._txByCat[cat].length); });
  if(!toFetch.length) return;
  // Fire calls sequentially to avoid hammering the Apps Script server.
  var i=0;
  function fetchNext(){
    if(i>=toFetch.length) return;
    var cat=toFetch[i++];
    google.script.run
      .withSuccessHandler(function(res){
        if(res && res.ok){ allData._txByCat[cat]=res.txs||[]; txByCatReady=true; }
        fetchNext();
      })
      .withFailureHandler(function(){ fetchNext(); })
      .getTxForCategory(cat, "all", 6);
  }
  fetchNext();
}

// ── Load data ─────────────────────────────────────────────────
function renderDashboard(data){
  data = data || {};
  allData = data;
  allData.allMonthlyTotals = data.allMonthlyTotals || [];
  allData.categoryTotals = data.categoryTotals || [];
  allData.monthCatData = data.monthCatData || {};
  allData.heatmapData = data.heatmapData || {months:[], cats:[], values:[]};
  allData.availableMonths = data.availableMonths || [];
  allData._txByCat = data.budgetTxByCat || data.txByCat || {};
  txByCatReady = Object.keys(allData._txByCat || {}).length > 0;

  const all = allData.allMonthlyTotals;
  const total = all.reduce((s, m) => s + (Number(m.total) || 0), 0);
  const avg = all.length ? total / all.length : 0;
  const curr = all.length ? (Number(all[all.length - 1].total) || 0) : 0;
  const totals = all.map(m => Number(m.total) || 0);
  const best = totals.length ? Math.min(...totals) : 0;
  const worst = totals.length ? Math.max(...totals) : 0;

  document.getElementById("kpi-total").textContent = fmtFull(total);
  document.getElementById("kpi-months").textContent = all.length;
  document.getElementById("kpi-avg").textContent = fmtFull(avg);
  document.getElementById("kpi-current").textContent = fmtFull(curr);
  document.getElementById("kpi-best").textContent = fmtFull(best);
  document.getElementById("kpi-worst").textContent = fmtFull(worst);

  _doUpdateRemaining();
  renderAllCharts();
  renderLayersPanel(data.layerData);
  renderBudgetPanel(data.budgetData);
  renderLeaksPanel(data.leaks);
  setTimeout(warmTxCacheForCurrentViews, 250);
}

function loadData(forceRefresh){
  var lw=parseInt((document.getElementById("leak-window")||{}).value||"3")||3;
  document.getElementById("loading").textContent="⏳ Loading data...";
  document.getElementById("loading").style.display="block";
  document.getElementById("app-inner").style.display="none";
  document.getElementById("error-msg").style.display="none";
  document.getElementById("cache-badge").style.display="none";
  txByCatReady=false;destroyAll();
  google.script.run
    .withSuccessHandler(function(res){
      document.getElementById("loading").style.display="none";
      if(!res||!res.ok||!res.data){document.getElementById("error-msg").textContent="❌ "+(res&&res.error||"No data.");document.getElementById("error-msg").style.display="block";return;}
      document.getElementById("app-inner").style.display="flex";
      if(res.fromCache)document.getElementById("cache-badge").style.display="inline";
      if(res.data.budgetTxByCat&&Object.keys(res.data.budgetTxByCat).length>0)txByCatReady=true;
      else if(res.data.txByCat&&Object.keys(res.data.txByCat).length>0)txByCatReady=true;
      renderDashboard(res.data);
    })
    .withFailureHandler(function(err){
      document.getElementById("loading").style.display="none";
      document.getElementById("error-msg").textContent="❌ "+(err&&err.message||String(err));
      document.getElementById("error-msg").style.display="block";
    })
    .getDashboardData(lw, forceRefresh===true);
}

loadData();
</script>
</body>
</html>`;
}