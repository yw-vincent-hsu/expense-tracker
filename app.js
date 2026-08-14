// ============================================================
// 家計簿 · 個人記帳 PWA
// 資料源：Google Sheets（OAuth 唯讀存取）
// ============================================================

const CONFIG = {
  CLIENT_ID: "267653972032-c7cc4oqq2fc96aob25uomgs41mqoej91.apps.googleusercontent.com",
  SPREADSHEET_ID: "1tiSbftHD85lhfrW1b792M217G-CbQ1A0KEdia1rn2Q0",
  RANGE: "A:F",
  SCOPES: "https://www.googleapis.com/auth/spreadsheets.readonly",
};

// Category color mapping — matcha family for expense, persimmon family for income
const EXPENSE_COLORS = {
  "餐飲": "#7A8F5C",
  "日用": "#A3B387",
  "育兒": "#566B32",
  "交通": "#C8D2B4",
  "醫療": "#8FA876",
  "娛樂": "#95A87A",
  "學習": "#B7C29E",
  "其他": "#D9DFC9",
};
const INCOME_COLORS = {
  "投資": "#C17A4D",
  "薪資": "#D4956B",
  "其他": "#E8C4A3",
};
function colorFor(cat, type){
  const map = type === "支出" ? EXPENSE_COLORS : INCOME_COLORS;
  if (map[cat]) return map[cat];
  // fallback deterministic color
  const fallback = type === "支出" ? ["#7A8F5C","#A3B387","#566B32","#8FA876"] : ["#C17A4D","#D4956B","#E8C4A3"];
  let hash = 0;
  for (const ch of cat) hash = (hash * 31 + ch.charCodeAt(0)) % fallback.length;
  return fallback[hash];
}

let tokenClient;
let accessToken = null;
let rawRows = [];        // parsed sheet rows
let currentMonth = null; // "2026-08"
let currentMode = "支出"; // pie mode: 支出 / 收入
let currentSort = "date";
let calSelectedDate = null;

// ---------- Boot ----------
window.onload = () => {
  bindTabs();
  bindSheet();
  initGoogleAuth();
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js").catch(() => {});
  }
};

function bindTabs(){
  document.querySelectorAll(".tab").forEach(tab => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
      tab.classList.add("active");
      const target = tab.dataset.view;
      document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
      document.getElementById("view-" + target).classList.add("active");
    });
  });
}

// ---------- Google OAuth ----------
function initGoogleAuth(){
  renderStatus("pieContent", "loading", "正在準備連線…");
  renderStatus("calendarContent", "loading", "正在準備連線…");

  const gsiScript = document.createElement("script");
  gsiScript.src = "https://accounts.google.com/gsi/client";
  gsiScript.onload = () => {
    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CONFIG.CLIENT_ID,
      scope: CONFIG.SCOPES,
      callback: (resp) => {
        if (resp.error) {
          renderStatus("pieContent", "error", "授權失敗，請重試", true);
          return;
        }
        accessToken = resp.access_token;
        fetchSheetData();
      },
    });
    promptSignIn();
  };
  document.head.appendChild(gsiScript);
}

function promptSignIn(){
  renderConnectPrompt();
}

function renderConnectPrompt(){
  const html = `
    <div class="status-msg">
      <span class="icon">🔒</span>
      連接你的 Google 帳號<br>以讀取「個人記帳本」試算表
      <br>
      <button class="btn" id="connectBtn">連接 Google 帳號</button>
    </div>`;
  document.getElementById("pieContent").innerHTML = html;
  document.getElementById("calendarContent").innerHTML = html.replace('id="connectBtn"', 'id="connectBtn2"');
  document.getElementById("connectBtn").addEventListener("click", () => tokenClient.requestAccessToken({ prompt: "" }));
  const btn2 = document.getElementById("connectBtn2");
  if (btn2) btn2.addEventListener("click", () => tokenClient.requestAccessToken({ prompt: "" }));
}

function renderStatus(elId, kind, msg, showRetry){
  const icon = kind === "error" ? "⚠️" : kind === "empty" ? "🗒️" : "◌";
  document.getElementById(elId).innerHTML = `
    <div class="status-msg">
      <span class="icon">${icon}</span>
      ${msg}
      ${showRetry ? '<br><button class="btn ghost" onclick="location.reload()">重新整理</button>' : ""}
    </div>`;
}

// ---------- Fetch & Parse Sheet ----------
async function fetchSheetData(){
  renderStatus("pieContent", "loading", "讀取記帳資料中…");
  try {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SPREADSHEET_ID}/values/${CONFIG.RANGE}`;
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (!res.ok) throw new Error("HTTP " + res.status);
    const data = await res.json();
    rawRows = parseRows(data.values || []);
    if (rawRows.length === 0) {
      renderStatus("pieContent", "empty", "目前沒有記帳資料");
      renderStatus("calendarContent", "empty", "目前沒有記帳資料");
      return;
    }
    setupMonthSelector();
    renderAll();
  } catch (err) {
    console.error(err);
    renderStatus("pieContent", "error", "讀取失敗，請確認網路連線或重新授權", true);
    renderStatus("calendarContent", "error", "讀取失敗，請確認網路連線或重新授權", true);
  }
}

function parseRows(values){
  if (values.length < 2) return [];
  const header = values[0];
  const idx = {
    date: header.indexOf("日期"),
    type: header.indexOf("類型"),
    category: header.indexOf("分類"),
    name: header.indexOf("項目名稱"),
    amount: header.indexOf("金額"),
    note: header.indexOf("備註"),
  };
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r || !r[idx.date] || !r[idx.type]) continue;
    const dateStr = String(r[idx.date]).trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) continue;
    const amountRaw = (r[idx.amount] || "0").toString().replace(/,/g, "");
    const amount = parseFloat(amountRaw) || 0;
    if (amount === 0) continue;
    rows.push({
      date: dateStr,
      month: dateStr.slice(0, 7),
      type: (r[idx.type] || "").trim(),
      category: (r[idx.category] || "其他").trim(),
      name: (r[idx.name] || "").trim(),
      amount: amount,
      note: (r[idx.note] || "").trim(),
    });
  }
  return rows;
}

// ---------- Month Selector ----------
function setupMonthSelector(){
  const months = [...new Set(rawRows.map(r => r.month))].sort().reverse();
  const sel = document.getElementById("monthSelect");
  sel.innerHTML = months.map(m => {
    const [y, mo] = m.split("-");
    return `<option value="${m}">${y}年${parseInt(mo)}月</option>`;
  }).join("");
  currentMonth = months[0];
  sel.value = currentMonth;
  sel.addEventListener("change", () => {
    currentMonth = sel.value;
    calSelectedDate = null;
    renderAll();
  });
}

function renderAll(){
  renderPieView();
  renderCalendarView();
}

// ================= PIE VIEW =================
function renderPieView(){
  const monthRows = rawRows.filter(r => r.month === currentMonth && r.type === currentMode);
  const container = document.getElementById("pieContent");

  if (monthRows.length === 0) {
    container.innerHTML = `
      <div class="mode-toggle" id="modeToggle">
        <div class="mode-btn ${currentMode==='支出'?'active-expense':''}" data-mode="支出">支出</div>
        <div class="mode-btn ${currentMode==='收入'?'active-income':''}" data-mode="收入">收入</div>
      </div>
      <div class="status-msg"><span class="icon">🗒️</span>這個月還沒有${currentMode}紀錄</div>`;
    bindModeToggle();
    return;
  }

  const byCat = {};
  let total = 0;
  monthRows.forEach(r => {
    byCat[r.category] = (byCat[r.category] || 0) + r.amount;
    total += r.amount;
  });
  const cats = Object.entries(byCat).sort((a, b) => b[1] - a[1]);

  const gradient = buildConicGradient(cats, total, currentMode);
  const themeColor = currentMode === "支出" ? "var(--matcha-deep)" : "var(--persimmon-deep)";

  let html = `
    <div class="mode-toggle" id="modeToggle">
      <div class="mode-btn ${currentMode==='支出'?'active-expense':''}" data-mode="支出">支出</div>
      <div class="mode-btn ${currentMode==='收入'?'active-income':''}" data-mode="收入">收入</div>
    </div>
    <div class="donut-card">
      <div class="donut-wrap">
        <svg width="220" height="220" viewBox="0 0 220 220">
          <circle cx="110" cy="110" r="80" fill="none" stroke="#EFE4CF" stroke-width="30"/>
          ${buildDonutSegments(cats, total)}
        </svg>
        <div class="donut-center">
          <div class="label">${currentMode}</div>
          <div class="amount" style="color:${themeColor}">${formatMoney(total)}</div>
        </div>
      </div>
      <div class="legend">
        ${cats.map(([cat, amt]) => {
          const pct = ((amt / total) * 100).toFixed(1);
          return `
          <div class="legend-row" data-cat="${escapeAttr(cat)}">
            <div class="legend-dot" style="background:${colorFor(cat, currentMode)}"></div>
            <div class="legend-name">${cat}</div>
            <div class="legend-pct">${pct}%</div>
            <div class="legend-amount">${formatMoney(amt)}</div>
            <div class="legend-chevron"></div>
          </div>`;
        }).join("")}
      </div>
    </div>
  `;
  container.innerHTML = html;
  bindModeToggle();

  document.querySelectorAll(".legend-row").forEach(row => {
    row.addEventListener("click", () => openCategorySheet(row.dataset.cat));
  });
}

function bindModeToggle(){
  document.querySelectorAll("#modeToggle .mode-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      currentMode = btn.dataset.mode;
      renderPieView();
    });
  });
}

function buildDonutSegments(cats, total){
  const r = 80, cx = 110, cy = 110, circumference = 2 * Math.PI * r;
  let offset = 0;
  return cats.map(([cat, amt]) => {
    const frac = amt / total;
    const len = frac * circumference;
    const seg = `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${colorFor(cat, currentMode)}"
      stroke-width="30" stroke-dasharray="${len} ${circumference - len}"
      stroke-dashoffset="${-offset}" transform="rotate(-90 ${cx} ${cy})" stroke-linecap="butt"/>`;
    offset += len;
    return seg;
  }).join("");
}

function buildConicGradient(){ return ""; } // unused, kept for potential fallback

// ---------- Category Detail Sheet ----------
function openCategorySheet(category){
  const rows = rawRows.filter(r => r.month === currentMonth && r.type === currentMode && r.category === category);
  const total = rows.reduce((s, r) => s + r.amount, 0);

  document.getElementById("sheetTitle").innerHTML =
    `<span style="width:11px;height:11px;border-radius:3px;display:inline-block;background:${colorFor(category, currentMode)}"></span> ${category}`;
  document.getElementById("sheetTotal").textContent = `${rows.length} 筆 · ${formatMoney(total)}`;

  currentSort = "date";
  document.querySelectorAll(".sort-btn").forEach(b => b.classList.toggle("active", b.dataset.sort === "date"));
  renderSheetList(rows);

  document.getElementById("sheetOverlay").classList.add("open");
  document.getElementById("sheet").classList.add("open");
}

function renderSheetList(rows){
  const sorted = [...rows].sort((a, b) => {
    if (currentSort === "amount") return b.amount - a.amount;
    return b.date.localeCompare(a.date);
  });
  document.getElementById("sheetList").innerHTML = sorted.map(r => `
    <div class="entry-row">
      <div class="entry-main">
        <div class="entry-name">${r.name || "（未命名）"}</div>
        <div class="entry-meta">${formatDateShort(r.date)}${r.note ? " · " + r.note : ""}</div>
      </div>
      <div class="entry-amount">${formatMoney(r.amount)}</div>
    </div>
  `).join("");
}

function bindSheet(){
  document.getElementById("sheetOverlay").addEventListener("click", closeSheet);
  document.querySelectorAll(".sort-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      currentSort = btn.dataset.sort;
      document.querySelectorAll(".sort-btn").forEach(b => b.classList.toggle("active", b === btn));
      const category = document.getElementById("sheetTitle").textContent.trim();
      const rows = rawRows.filter(r => r.month === currentMonth && r.type === currentMode && r.category === category);
      renderSheetList(rows);
    });
  });
}
function closeSheet(){
  document.getElementById("sheetOverlay").classList.remove("open");
  document.getElementById("sheet").classList.remove("open");
}

// ================= CALENDAR VIEW =================
function renderCalendarView(){
  const monthRows = rawRows.filter(r => r.month === currentMonth);
  const container = document.getElementById("calendarContent");

  const totalExpense = monthRows.filter(r => r.type === "支出").reduce((s, r) => s + r.amount, 0);
  const totalIncome = monthRows.filter(r => r.type === "收入").reduce((s, r) => s + r.amount, 0);

  const byDay = {};
  monthRows.forEach(r => {
    if (!byDay[r.date]) byDay[r.date] = { expense: 0, income: 0, rows: [] };
    if (r.type === "支出") byDay[r.date].expense += r.amount;
    else byDay[r.date].income += r.amount;
    byDay[r.date].rows.push(r);
  });

  const [y, m] = currentMonth.split("-").map(Number);
  const firstDay = new Date(y, m - 1, 1);
  const daysInMonth = new Date(y, m, 0).getDate();
  const startWeekday = (firstDay.getDay() + 6) % 7; // Monday = 0

  const todayStr = new Date().toISOString().slice(0, 10);

  let cells = "";
  for (let i = 0; i < startWeekday; i++) cells += `<div class="cal-day empty"></div>`;
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${currentMonth}-${String(d).padStart(2, "0")}`;
    const info = byDay[dateStr];
    const isToday = dateStr === todayStr;
    const isSelected = dateStr === calSelectedDate;
    let amtHtml = "";
    if (info && info.expense > 0) amtHtml = `<div class="cal-day-amt">-${formatCompact(info.expense)}</div>`;
    else if (info && info.income > 0) amtHtml = `<div class="cal-day-amt" style="color:var(--persimmon-deep)">+${formatCompact(info.income)}</div>`;
    cells += `
      <div class="cal-day ${info ? "has-data" : ""} ${isToday ? "today" : ""} ${isSelected ? "selected" : ""}" data-date="${dateStr}">
        <div class="cal-day-num">${d}</div>
        ${amtHtml}
      </div>`;
  }

  let html = `
    <div class="cal-summary">
      <div class="cal-summary-item">
        <div class="cal-summary-label">總支出</div>
        <div class="cal-summary-value expense">${formatMoney(totalExpense)}</div>
      </div>
      <div class="cal-summary-divider"></div>
      <div class="cal-summary-item">
        <div class="cal-summary-label">總收入</div>
        <div class="cal-summary-value income">${formatMoney(totalIncome)}</div>
      </div>
    </div>
    <div class="cal-grid-card">
      <div class="cal-weekdays"><div>一</div><div>二</div><div>三</div><div>四</div><div>五</div><div>六</div><div>日</div></div>
      <div class="cal-days">${cells}</div>
    </div>
    <div id="dayDetail"></div>
  `;
  container.innerHTML = html;

  document.querySelectorAll(".cal-day[data-date]").forEach(cell => {
    cell.addEventListener("click", () => {
      calSelectedDate = cell.dataset.date;
      renderCalendarView();
      document.getElementById("dayDetail").scrollIntoView({ behavior: "smooth", block: "nearest" });
    });
  });

  if (calSelectedDate && byDay[calSelectedDate]) {
    renderDayDetail(calSelectedDate, byDay[calSelectedDate]);
  } else if (calSelectedDate) {
    document.getElementById("dayDetail").innerHTML = `
      <div class="day-detail-card"><div class="status-msg" style="padding:32px 20px;">這天沒有紀錄</div></div>`;
  }
}

function renderDayDetail(dateStr, info){
  const net = info.income - info.expense;
  const rows = [...info.rows].sort((a, b) => b.amount - a.amount);
  const html = `
    <div class="day-detail-card">
      <div class="day-detail-header">
        <div class="day-detail-date">${formatDateFull(dateStr)}</div>
        <div class="day-detail-net" style="color:${net >= 0 ? 'var(--persimmon-deep)' : 'var(--matcha-deep)'}">
          ${net >= 0 ? "+" : ""}${formatMoney(net)}
        </div>
      </div>
      <div class="day-detail-list">
        ${rows.map(r => `
          <div class="entry-row">
            <div class="entry-main">
              <div class="entry-name">${r.name || "（未命名）"}</div>
              <div class="entry-meta">${r.category}${r.note ? " · " + r.note : ""}</div>
            </div>
            <div class="entry-amount" style="color:${r.type==='支出' ? 'var(--ink)' : 'var(--persimmon-deep)'}">
              ${r.type === "支出" ? "-" : "+"}${formatMoney(r.amount)}
            </div>
          </div>
        `).join("")}
      </div>
    </div>
  `;
  document.getElementById("dayDetail").innerHTML = html;
}

// ---------- Formatters ----------
function formatMoney(n){
  return Math.round(n).toLocaleString("zh-TW");
}
function formatCompact(n){
  if (n >= 1000) return (n / 1000).toFixed(1).replace(/\.0$/, "") + "k";
  return String(Math.round(n));
}
function formatDateShort(dateStr){
  const [, m, d] = dateStr.split("-");
  const weekday = "日一二三四五六"[new Date(dateStr).getDay()];
  return `${parseInt(m)}/${parseInt(d)}（${weekday}）`;
}
function formatDateFull(dateStr){
  const [y, m, d] = dateStr.split("-");
  const weekday = "日一二三四五六"[new Date(dateStr).getDay()];
  return `${parseInt(m)}月${parseInt(d)}日（${weekday}）`;
}
function escapeAttr(s){
  return String(s).replace(/"/g, "&quot;");
}
