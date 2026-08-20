// ============================================================
// 家計簿 · 個人記帳 PWA
// 資料源：Google Sheets（OAuth 讀寫存取）
// ============================================================

const CONFIG = {
  CLIENT_ID: "267653972032-c7cc4oqq2fc96aob25uomgs41mqoej91.apps.googleusercontent.com",
  SPREADSHEET_ID: "1tiSbftHD85lhfrW1b792M217G-CbQ1A0KEdia1rn2Q0",
  RANGE: "A:F",
  SCOPES: "https://www.googleapis.com/auth/spreadsheets",
};

// Color palettes — assigned dynamically by rank (largest category gets the
// first/deepest color), not tied to any fixed category name. This means the
// ring naturally spreads distinct tones across adjacent slices without needing
// a fixed per-category mapping, since whichever categories are biggest this
// month simply take the first colors in the list.
const EXPENSE_PALETTE = [
  "#6A8372", // 老竹
  "#465D4C", // 御納戸茶
  "#36563C", // 千歳緑
  "#516E41", // 青丹
  "#7BA23F", // 萌黃
  "#90B44B", // 鶸萌黃
  "#BEC23F", // 鶸
  "#DDD23B", // 女郎花
];
const INCOME_PALETTE = [
  "#B47157", // 唐茶
  "#E79460", // 洗柿
  "#FFBA84", // 灑落柿
];

function paletteFor(type){
  return type === "支出" ? EXPENSE_PALETTE : INCOME_PALETTE;
}
// Deterministic fallback for any category beyond the palette length (e.g. a
// 9th expense category some month) — cycles the palette rather than repeating
// the last color indefinitely, keeps things visually distinguishable.
function colorForRank(rank, type){
  const palette = paletteFor(type);
  return palette[rank % palette.length];
}

let tokenClient;
let accessToken = null;
let rawRows = [];        // parsed sheet rows
let currentMonth = null; // "2026-08"
let currentMode = "支出"; // pie mode: 支出 / 收入
let currentSort = "date";
let currentSortDir = "desc"; // "desc" or "asc" — toggled by re-clicking the active sort-btn
let calSelectedDate = null;
let pollTimer = null;
let lastSheetSignature = null; // used to detect real changes silently
let pendingEntry = null;
const ENTRY_CATEGORIES = {
  "支出": ["餐飲", "交通", "日用", "娛樂", "育兒", "醫療", "學習", "其他"],
  "收入": ["薪資", "投資", "其它"],
};

const TOKEN_STORAGE_KEY = "kakeibo_token";
const POLL_INTERVAL_MS = 60 * 1000; // check for sheet updates every 60s while app is open

// ---------- Token persistence ----------
function saveToken(token, expiresInSec){
  const expiresAt = Date.now() + expiresInSec * 1000;
  localStorage.setItem(TOKEN_STORAGE_KEY, JSON.stringify({ token, expiresAt }));
}
function loadStoredToken(){
  try {
    const raw = localStorage.getItem(TOKEN_STORAGE_KEY);
    if (!raw) return null;
    const { token, expiresAt } = JSON.parse(raw);
    // leave a 60s safety margin before actual expiry
    if (Date.now() > expiresAt - 60 * 1000) {
      localStorage.removeItem(TOKEN_STORAGE_KEY);
      return null;
    }
    return token;
  } catch {
    return null;
  }
}
function clearStoredToken(){
  localStorage.removeItem(TOKEN_STORAGE_KEY);
}

// ---------- Boot ----------
window.onload = () => {
  bindTabs();
  bindSheet();
  bindEntryForm();
  initGoogleAuth();
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js").catch(() => {});
  }
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible" && accessToken) {
      fetchSheetData(true); // silent refresh when user comes back to the tab
    }
  });
};

function bindTabs(){
  const indicator = document.getElementById("tabsIndicator");
  const moveIndicator = (tabEl) => {
    indicator.style.width = tabEl.offsetWidth + "px";
    indicator.style.transform = `translateX(${tabEl.offsetLeft}px)`;
  };
  document.querySelectorAll(".tab").forEach(tab => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
      tab.classList.add("active");
      moveIndicator(tab);
      const target = tab.dataset.view;
      if (target === "pie") {
        closeSheet();
        if (!pendingEntry) closeEntrySheet();
        currentMode = "支出";
        calSelectedDate = null;
        resetPageScroll();
        if (document.getElementById("pieBody")) {
          updateModeToggleUI();
          renderPieBody();
        }
      }
      if (target === "calendar") resetPageScroll();
      document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
      document.getElementById("view-" + target).classList.add("active");
      // The 圖表 tab is laid out to fit one screen with no scrolling; the
      // 日曆 tab still needs to scroll for long day-detail lists.
      document.body.classList.toggle("pie-locked", target === "pie");
    });
  });
  moveIndicator(document.querySelector(".tab.active"));
  document.body.classList.toggle("pie-locked", document.querySelector(".tab.active").dataset.view === "pie");
  window.addEventListener("resize", () => moveIndicator(document.querySelector(".tab.active")));
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
        saveToken(resp.access_token, resp.expires_in || 3600);
        if (pendingEntry) {
          const entry = pendingEntry;
          pendingEntry = null;
          submitEntry(entry);
        } else {
          fetchSheetData();
          startPolling();
        }
      },
    });

    // Try to reuse a still-valid token from a previous session first —
    // this is what avoids re-prompting on every reload within the ~1hr window.
    const stored = loadStoredToken();
    if (stored) {
      accessToken = stored;
      fetchSheetData();
      startPolling();
    } else {
      promptSignIn();
    }
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
// silent=true is used for background polling / visibility-refresh: it never
// shows a loading spinner and never wipes the screen on failure, so a flaky
// network blip in the background doesn't disrupt whatever the user is looking at.
async function fetchSheetData(silent){
  if (!silent) {
    renderStatus("pieContent", "loading", "讀取記帳資料中…");
  }
  try {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SPREADSHEET_ID}/values/${CONFIG.RANGE}`;
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (res.status === 401) {
      // token expired or was revoked — drop it and ask the user to sign in again.
      // Even in silent mode we still surface the prompt here (rather than staying
      // quiet), because a session that's silently gone stale with no way for the
      // user to notice is worse than one unprompted screen change.
      clearStoredToken();
      accessToken = null;
      stopPolling();
      renderConnectPrompt();
      return;
    }
    if (!res.ok) throw new Error("HTTP " + res.status);
    const data = await res.json();
    const values = data.values || [];

    // Cheap signature to detect whether anything actually changed before
    // re-rendering — avoids the UI silently jumping/resetting scroll position
    // every 60s when Gemini Spark hasn't written anything new.
    const signature = JSON.stringify(values);
    const changed = signature !== lastSheetSignature;
    lastSheetSignature = signature;

    if (!changed && silent) return;

    rawRows = parseRows(values);
    if (rawRows.length === 0) {
      renderStatus("pieContent", "empty", "目前沒有記帳資料");
      renderStatus("calendarContent", "empty", "目前沒有記帳資料");
      return;
    }
    setupMonthSelector(currentMonth);
    renderAll();

    // if the category detail sheet is open, refresh it too so it doesn't show stale numbers
    const sheetEl = document.getElementById("sheet");
    if (sheetEl.classList.contains("open")) {
      const category = document.getElementById("sheetTitle").textContent.trim();
      const rows = rawRows.filter(r => r.month === currentMonth && r.type === currentMode && r.category === category);
      renderSheetList(rows);
    }
  } catch (err) {
    console.error(err);
    if (!silent) {
      renderStatus("pieContent", "error", "讀取失敗，請確認網路連線或重新授權", true);
      renderStatus("calendarContent", "error", "讀取失敗，請確認網路連線或重新授權", true);
    }
  }
}

// ---------- Silent background polling ----------
// Checks Google Sheets periodically while the tab is open/visible so entries
// added via Gemini Spark show up without the user having to manually reload.
function startPolling(){
  stopPolling();
  pollTimer = setInterval(() => {
    if (document.visibilityState === "visible" && accessToken) {
      fetchSheetData(true);
    }
  }, POLL_INTERVAL_MS);
}
function stopPolling(){
  if (pollTimer) clearInterval(pollTimer);
  pollTimer = null;
}

// ================= ADD ENTRY =================
function bindEntryForm(){
  const form = document.getElementById("entryForm");
  resetEntryForm();

  document.getElementById("addEntryBtn").addEventListener("click", openEntrySheet);
  document.getElementById("cancelEntryBtn").addEventListener("click", closeEntrySheet);
  document.getElementById("entryOverlay").addEventListener("click", closeEntrySheet);
  document.querySelectorAll("[data-entry-type]").forEach(btn => {
    btn.addEventListener("click", () => {
      document.getElementById("entryType").value = btn.dataset.entryType;
      document.querySelectorAll("[data-entry-type]").forEach(typeBtn => {
        typeBtn.classList.toggle("active-expense", typeBtn.dataset.entryType === "支出" && typeBtn === btn);
        typeBtn.classList.toggle("active-income", typeBtn.dataset.entryType === "收入" && typeBtn === btn);
      });
      updateEntryCategories();
    });
  });
  form.addEventListener("submit", (event) => {
    event.preventDefault();
    const displayDate = document.getElementById("entryDate").value.trim();
    const date = parseDisplayDate(displayDate);
    const entry = {
      date,
      type: document.getElementById("entryType").value,
      category: document.getElementById("entryCategory").value.trim(),
      name: document.getElementById("entryName").value.trim(),
      amount: Number(document.getElementById("entryAmount").value),
      note: document.getElementById("entryNote").value.trim(),
    };
    if (!entry.date || !entry.category || !entry.name || !Number.isFinite(entry.amount) || entry.amount <= 0) {
      document.getElementById("entryError").textContent = "請填寫正確日期（例如 2026/8/21）、分類、項目名稱與有效金額";
      return;
    }
    submitEntry(entry);
  });
}

function openEntrySheet(){
  resetEntryForm();
  document.getElementById("entryError").textContent = "";
  document.getElementById("entryOverlay").classList.add("open");
  document.getElementById("entrySheet").classList.add("open");
}

function resetEntryForm(){
  const form = document.getElementById("entryForm");
  if (form) form.reset();
  const today = new Date();
  document.getElementById("entryDate").value = `${today.getFullYear()}/${today.getMonth() + 1}/${today.getDate()}`;
  document.getElementById("entryType").value = "支出";
  document.querySelectorAll("[data-entry-type]").forEach(btn => {
    btn.classList.toggle("active-expense", btn.dataset.entryType === "支出");
    btn.classList.toggle("active-income", false);
  });
  updateEntryCategories();
}

function updateEntryCategories(){
  const type = document.getElementById("entryType").value;
  const select = document.getElementById("entryCategory");
  select.innerHTML = `<option value="" disabled selected>請選擇分類</option>${ENTRY_CATEGORIES[type].map(category => `<option value="${category}">${category}</option>`).join("")}`;
}

function parseDisplayDate(value){
  const match = /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/.exec(value);
  if (!match) return "";
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return "";
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function closeEntrySheet(){
  if (pendingEntry) return;
  if (document.activeElement && typeof document.activeElement.blur === "function") {
    document.activeElement.blur();
  }
  document.getElementById("entryOverlay").classList.remove("open");
  document.getElementById("entrySheet").classList.remove("open");
  resetPageScroll();
}

function resetPageScroll(){
  const restore = () => {
    window.scrollTo(0, 0);
    document.documentElement.scrollTop = 0;
    document.body.scrollTop = 0;
  };
  restore();
  requestAnimationFrame(() => {
    restore();
    requestAnimationFrame(restore);
  });
  setTimeout(restore, 250);
}

async function submitEntry(entry){
  const errorEl = document.getElementById("entryError");
  const saveBtn = document.getElementById("saveEntryBtn");
  if (!accessToken) {
    pendingEntry = entry;
    errorEl.textContent = "請先連接 Google 帳號，授權後會繼續儲存";
    promptSignIn();
    return;
  }
  saveBtn.disabled = true;
  saveBtn.textContent = "儲存中…";
  errorEl.textContent = "";
  try {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SPREADSHEET_ID}/values/${CONFIG.RANGE}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`;
    const res = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ values: [[entry.date, entry.type, entry.category, entry.name, entry.amount, entry.note]] }),
    });
    if (res.status === 401) {
      clearStoredToken();
      accessToken = null;
      pendingEntry = entry;
      promptSignIn();
      throw new Error("授權已過期");
    }
    if (res.status === 403) {
      pendingEntry = entry;
      errorEl.textContent = "需要新增 Google 試算表的寫入權限，請重新授權";
      tokenClient.requestAccessToken({ prompt: "consent" });
      return;
    }
    if (!res.ok) throw new Error("HTTP " + res.status);
    closeEntrySheet();
    resetEntryForm();
    lastSheetSignature = null;
    await fetchSheetData(true);
  } catch (err) {
    if (err.message !== "授權已過期") {
      console.error(err);
      errorEl.textContent = "儲存失敗，請確認網路連線後再試一次";
    }
  } finally {
    saveBtn.disabled = false;
    saveBtn.textContent = "儲存記帳";
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
// preferredMonth: pass the currently-viewed month (e.g. during a silent
// background refresh) so polling doesn't yank the user back to the latest
// month while they're looking at an older one. Pass null on first load to
// default to the most recent month.
function setupMonthSelector(preferredMonth){
  const months = [...new Set(rawRows.map(r => r.month))].sort().reverse();
  const sel = document.getElementById("monthSelect");
  const isFirstBuild = !sel.dataset.bound;

  sel.innerHTML = months.map(m => {
    const [y, mo] = m.split("-");
    return `<option value="${m}">${y}年${parseInt(mo)}月</option>`;
  }).join("");

  currentMonth = (preferredMonth && months.includes(preferredMonth)) ? preferredMonth : months[0];
  sel.value = currentMonth;

  if (isFirstBuild) {
    sel.dataset.bound = "1";
    sel.addEventListener("change", () => {
      currentMonth = sel.value;
      calSelectedDate = null;
      renderAll();
    });
  }
}

function renderAll(){
  renderPieView();
  renderCalendarView();
}

// ================= PIE VIEW =================
function renderPieView(){
  const container = document.getElementById("pieContent");
  if (!document.getElementById("modeToggle")) {
    container.innerHTML = `
      <div class="mode-toggle" id="modeToggle">
        <div class="mode-toggle-indicator" id="modeIndicator"></div>
        <div class="mode-btn" data-mode="支出">支出</div>
        <div class="mode-btn" data-mode="收入">收入</div>
      </div>
      <div id="pieBody"></div>
    `;
    bindModeToggle();
  }
  updateModeToggleUI();
  renderPieBody();
}

function updateModeToggleUI(){
  document.querySelectorAll("#modeToggle .mode-btn").forEach(btn => {
    const isActive = btn.dataset.mode === currentMode;
    btn.classList.toggle("active-expense", isActive && currentMode === "支出");
    btn.classList.toggle("active-income", isActive && currentMode === "收入");
  });
  document.getElementById("modeIndicator").classList.toggle("income", currentMode === "收入");
}

// Fades the donut/legend body out, repaints it for the current month+mode,
// then fades it back in — skipped on the very first paint so initial load
// isn't delayed waiting on a transition with nothing to cross-fade from.
function renderPieBody(){
  const body = document.getElementById("pieBody");
  const isFirstPaint = !body.innerHTML.trim();

  const paint = () => {
    const monthRows = rawRows.filter(r => r.month === currentMonth && r.type === currentMode);

    if (monthRows.length === 0) {
      body.innerHTML = `<div class="status-msg"><span class="icon">🗒️</span>這個月還沒有${currentMode}紀錄</div>`;
      return;
    }

    const byCat = {};
    let total = 0;
    monthRows.forEach(r => {
      byCat[r.category] = (byCat[r.category] || 0) + r.amount;
      total += r.amount;
    });

    // Single amount-descending order drives both the ring and the legend now —
    // color is assigned by rank (biggest category = first/deepest palette color),
    // so ring and legend always agree on which color means which category.
    const sortedCats = Object.entries(byCat).sort((a, b) => b[1] - a[1]);
    const themeColor = colorForRank(0, currentMode); // center number matches the largest category's color

    body.innerHTML = `
      <div class="donut-card">
        <div class="donut-wrap">
          <svg width="100%" height="100%" viewBox="0 0 220 220">
            ${buildDonutSegments(sortedCats, total)}
          </svg>
          <div class="donut-center">
            <div class="label">${currentMode}</div>
            <div class="amount" style="color:${themeColor}">${formatMoney(total)}</div>
          </div>
        </div>
        <div class="legend">
          ${sortedCats.map(([cat, amt], i) => {
            const pct = ((amt / total) * 100).toFixed(1);
            return `
            <div class="legend-row" data-cat="${escapeAttr(cat)}" data-rank="${i}">
              <div class="legend-dot" style="background:${colorForRank(i, currentMode)}"></div>
              <div class="legend-label">
                <div class="legend-name">${cat}</div>
                <div class="legend-pct">${pct}%</div>
              </div>
              <div class="legend-amount">${formatMoney(amt)}</div>
              <div class="legend-chevron"></div>
            </div>`;
          }).join("")}
        </div>
      </div>
    `;
    document.querySelectorAll(".legend-row").forEach(row => {
      row.addEventListener("click", () => openCategorySheet(row.dataset.cat, parseInt(row.dataset.rank, 10)));
    });
  };

  if (isFirstPaint) {
    paint();
    return;
  }
  body.classList.add("fade-out");
  setTimeout(() => {
    paint();
    body.classList.remove("fade-out");
  }, 150);
}

function bindModeToggle(){
  document.querySelectorAll("#modeToggle .mode-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      if (btn.dataset.mode === currentMode) return;
      currentMode = btn.dataset.mode;
      updateModeToggleUI();
      renderPieBody();
    });
  });
}

// Draws the donut ring with segments touching edge-to-edge (no gap) — a
// small gap looked fine for medium-sized slices but broke down for very
// small categories (their slice could end up thinner than the gap itself,
// making them vanish or look like a stray sliver). Color is assigned by
// rank within this render (index 0 = biggest = first palette color).
function buildDonutSegments(cats, total){
  const r = 80, cx = 110, cy = 110, circumference = 2 * Math.PI * r;
  let offset = 0;
  return cats.map(([cat, amt], i) => {
    const frac = amt / total;
    const len = frac * circumference;
    const seg = `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${colorForRank(i, currentMode)}"
      stroke-width="30" stroke-dasharray="${len} ${circumference - len}"
      stroke-dashoffset="${-offset}" transform="rotate(-90 ${cx} ${cy})" stroke-linecap="butt"/>`;
    offset += len;
    return seg;
  }).join("");
}

// ---------- Category Detail Sheet ----------
// rank identifies this category's position in the current month's
// amount-sorted list, so the dot color matches what's shown in the ring/legend.
function openCategorySheet(category, rank){
  const rows = rawRows.filter(r => r.month === currentMonth && r.type === currentMode && r.category === category);
  const total = rows.reduce((s, r) => s + r.amount, 0);

  document.getElementById("sheetTitle").innerHTML =
    `<span style="width:11px;height:11px;border-radius:3px;display:inline-block;background:${colorForRank(rank, currentMode)}"></span> ${category}`;
  document.getElementById("sheetTotal").textContent = `${rows.length} 筆 · ${formatMoney(total)}`;
  document.querySelector(".sort-toggle").classList.toggle("income", currentMode === "收入");

  currentSort = "date";
  currentSortDir = "desc";
  updateSortButtons();
  renderSheetList(rows);

  document.getElementById("sheetOverlay").classList.add("open");
  document.getElementById("sheet").classList.add("open");
}

// Shows which sort is active and its direction (re-clicking the active
// button flips this instead of doing nothing).
function updateSortButtons(){
  const arrow = currentSortDir === "asc" ? "↑" : "↓";
  document.querySelectorAll(".sort-btn").forEach(btn => {
    const isActive = btn.dataset.sort === currentSort;
    btn.classList.toggle("active", isActive);
    const label = btn.dataset.sort === "date" ? "依日期" : "依金額";
    btn.textContent = `${label} ${isActive ? arrow : "↓"}`;
  });
}

function renderSheetList(rows){
  const dirMultiplier = currentSortDir === "asc" ? -1 : 1;
  const sorted = [...rows].sort((a, b) => {
    const cmp = currentSort === "amount" ? (b.amount - a.amount) : b.date.localeCompare(a.date);
    return cmp * dirMultiplier;
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
      if (btn.dataset.sort === currentSort) {
        currentSortDir = currentSortDir === "asc" ? "desc" : "asc";
      } else {
        currentSort = btn.dataset.sort;
        currentSortDir = "desc";
      }
      updateSortButtons();
      const category = document.getElementById("sheetTitle").textContent.trim();
      const rows = rawRows.filter(r => r.month === currentMonth && r.type === currentMode && r.category === category);
      renderSheetList(rows);
    });
  });
  bindSheetDrag();
  bindSheetDrag("entrySheet", "entryDragHandle", closeEntrySheet);
}
function closeSheet(){
  document.getElementById("sheetOverlay").classList.remove("open");
  document.getElementById("sheet").classList.remove("open");
}

// Lets the user pull the sheet down by its handle to dismiss it, mirroring
// native iOS bottom-sheet behavior. Pointer Events unify touch and mouse so
// the same code path works for on-device testing in a desktop browser too.
function bindSheetDrag(sheetId = "sheet", handleId = "sheetDragHandle", dismiss = closeSheet){
  const sheet = document.getElementById(sheetId);
  const dragHandle = document.getElementById(handleId);
  let startY = 0;
  let dragY = 0;
  let dragging = false;

  dragHandle.addEventListener("pointerdown", (e) => {
    dragging = true;
    startY = e.clientY;
    dragY = 0;
    sheet.style.transition = "none";
    dragHandle.setPointerCapture(e.pointerId);
  });
  dragHandle.addEventListener("pointermove", (e) => {
    if (!dragging) return;
    dragY = Math.max(0, e.clientY - startY);
    sheet.style.transform = `translateY(${dragY}px)`;
  });
  const endDrag = () => {
    if (!dragging) return;
    dragging = false;
    sheet.style.transition = "";
    sheet.style.transform = "";
    // Dismiss once dragged past a quarter of the sheet's own height (or a
    // flat 120px for very short sheets) — otherwise snap back open.
    if (dragY > sheet.offsetHeight * 0.25 || dragY > 120) {
      dismiss();
    }
  };
  dragHandle.addEventListener("pointerup", endDrag);
  dragHandle.addEventListener("pointercancel", endDrag);
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

  // Local calendar date, not UTC — toISOString() reports UTC, which still
  // reads as "yesterday" in UTC+8 until 08:00 local time and made the
  // today-marker fall a day behind.
  const now = new Date();
  const todayStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;

  let cells = "";
  for (let i = 0; i < startWeekday; i++) cells += `<div class="cal-day empty"></div>`;
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${currentMonth}-${String(d).padStart(2, "0")}`;
    const info = byDay[dateStr];
    const isToday = dateStr === todayStr;
    const isSelected = dateStr === calSelectedDate;
    let amtHtml = "";
    if (info) {
      const net = info.income - info.expense;
      if (net > 0) amtHtml = `<div class="cal-day-amt positive">+${formatCompact(net)}</div>`;
      else if (net < 0) amtHtml = `<div class="cal-day-amt">-${formatCompact(-net)}</div>`;
    }
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
        <div class="day-detail-net" style="color:${net >= 0 ? 'var(--income-accent)' : 'var(--expense-accent)'}">
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
            <div class="entry-amount" style="color:${r.type==='支出' ? 'var(--ink)' : 'var(--income-accent)'}">
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
  return `${parseInt(m)}/${parseInt(d)} (${weekday})`;
}
function formatDateFull(dateStr){
  const [y, m, d] = dateStr.split("-");
  const weekday = "日一二三四五六"[new Date(dateStr).getDay()];
  return `${parseInt(m)}/${parseInt(d)} (${weekday})`;
}
function escapeAttr(s){
  return String(s).replace(/"/g, "&quot;");
}
