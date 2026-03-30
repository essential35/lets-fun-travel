/* ═══════════════════════════════════════════
   旅途筆記 app.js  v2.1
   修改：
   1. Google Sheet 讀寫同步（透過 Apps Script Web App）
   2. 每個 Sheet URL 對應獨立行程（以 Sheet ID 為 localStorage key）
   3. 搜尋景點時在 modal 內的 iframe 預覽地圖
   ═══════════════════════════════════════════ */

// ── UTILITY（必須放最前面，避免 const 無 hoisting 導致函式找不到）──
const $ = id => document.getElementById(id);
const v = id => $(id).value;
const el = (tag, cls) => { const e = document.createElement(tag); e.className = cls; return e; };
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

function fmtShort(d) { return d ? (d.slice(2).replace(/-/g, '.')) : '-'; }
function fmtDT(dt) { return dt ? (dt.replace('T', ' ').slice(0, 16)) : '—'; }
function todayStr() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
function nowStr() { return new Date().toLocaleString('zh-TW', { hour12: false }).slice(0, 16); }
function addDays(d, n) {
  const [y, m, day] = d.split('-').map(Number);
  const dt = new Date(y, m - 1, day + n);
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`;
}
function diffDays(a, b) {
  const [ay, am, ad] = a.split('-').map(Number);
  const [by, bm, bd] = b.split('-').map(Number);
  return Math.round((new Date(by, bm - 1, bd) - new Date(ay, am - 1, ad)) / 86400000);
}
function blob(content, type, name) {
  const url = URL.createObjectURL(new Blob([content], { type }));
  const a = document.createElement('a'); a.href = url; a.download = name; a.click(); URL.revokeObjectURL(url);
}
function log(a) { S.history.unshift({ t: nowStr(), a }); }
function toast(msg) {
  const t = $('toast'); t.textContent = msg; t.classList.add('show');
  clearTimeout(t._t); t._t = setTimeout(() => t.classList.remove('show'), 2800);
}

// ── STATE ─────────────────────────────────
const S = {
  trip: { theme: '', start: '', end: '' },
  dataSource: '',
  sheetId: '',
  gasUrl: '',
  inviteCode: '',         // 驗證通過的邀請碼（每次 GAS 請求都帶上）
  isWhitelisted: false,   // GAS 驗證通過才為 true，否則僅本機模式
  selectedDate: null,
  itinerary: [],
  flights: [],
  accommodation: [],
  expenses: [],
  memos: [],
  history: [],
  pendingSpot: null,
  dragSrc: null,
  dragTarget: null,
  selectedItemId: null,
  showAllDays: false,
  currentMapUrl: '',
  syncTimer: null,
  lastSynced: null,
};

// ── INVITE CODE AUTH ──────────────────────
// 邀請碼整合在 setup modal 中，initTrip 時驗證
// 通過：isWhitelisted=true，同步正常運作
// 不通過或無 GAS：isWhitelisted=false，僅本機儲存

async function verifyInviteCode(gasUrl, sheetId, code) {
  if (!gasUrl) return { ok: false, reason: 'no-gas' };
  if (!code) return { ok: false, reason: 'no-code' };
  try {
    const idParam = sheetId ? `&id=${sheetId}` : '';
    const resp = await fetch(
      `${gasUrl}?action=auth&code=${encodeURIComponent(code)}${idParam}`
    );
    const data = await resp.json();
    return { ok: !!data.ok, reason: data.ok ? 'ok' : 'wrong-code' };
  } catch (e) {
    return { ok: false, reason: 'network-error' };
  }
}

function showInviteCodeStatus(msg, type = 'ok') {
  const el = document.getElementById('invite-code-status');
  if (!el) return;
  el.textContent = msg;
  el.style.display = 'block';
  el.style.background = type === 'ok' ? 'rgba(42,122,122,0.1)' :
    type === 'warn' ? 'rgba(184,145,58,0.1)' :
      type === 'info' ? 'rgba(42,122,122,0.06)' :
        'rgba(201,79,42,0.1)';
  el.style.color = type === 'ok' ? 'var(--teal)' :
    type === 'warn' ? 'var(--gold)' :
      type === 'info' ? 'var(--teal)' :
        'var(--accent)';
}

function updateAuthBadgeInApp() {
  let bar = document.getElementById('sync-user-badge');
  if (!bar) {
    bar = document.createElement('span');
    bar.id = 'sync-user-badge';
    bar.style.cssText = 'font-size:11px;color:var(--ink-3);flex-shrink:0;white-space:nowrap;';
    const syncBar = document.getElementById('sync-bar');
    if (syncBar) syncBar.appendChild(bar);
  }
  bar.textContent = S.isWhitelisted ? '🔑 已驗證' : '🔑 本機';
}

// 頁面載入時自動填入上次記住的邀請碼
window.addEventListener('DOMContentLoaded', () => {
  try {
    const savedCode = localStorage.getItem('travel_invite_code');
    if (savedCode) {
      const input = document.getElementById('invite-code-input');
      if (input) input.value = savedCode;
    }
    const savedGas = localStorage.getItem('travel_gas_url');
    if (savedGas) {
      const gasInput = document.getElementById('gas-url');
      if (gasInput && !gasInput.value) gasInput.value = savedGas;
    }
  } catch (e) { }
});

// ── GOOGLE SHEET ID 解析 ───────────────────
function extractSheetId(url) {
  if (!url) return '';
  const m = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : '';
}

// ── LOCAL STORAGE KEY ─────────────────────
// 以 sheetId 為 key，確保不同 Sheet URL 的用戶資料互相獨立
function storageKey() {
  return S.sheetId ? `travel_planner_${S.sheetId}` : 'travel_planner_default';
}

function saveToLocal() {
  try {
    const data = {
      trip: S.trip,
      dataSource: S.dataSource,
      gasUrl: S.gasUrl,
      itinerary: S.itinerary,
      flights: S.flights,
      accommodation: S.accommodation,
      expenses: S.expenses,
      memos: S.memos,
      history: S.history,
    };
    localStorage.setItem(storageKey(), JSON.stringify(data));
  } catch (e) { console.warn('localStorage save failed', e); }
}

function loadFromLocal(sheetId) {
  try {
    const key = sheetId ? `travel_planner_${sheetId}` : 'travel_planner_default';
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}

// ── GOOGLE APPS SCRIPT 同步 ───────────────────────────
// 需要用戶部署一個 GAS Web App，並將 Web App URL 填入設定中
// GAS 程式碼範本（貼入 script.google.com）：
// ----------------------------------------
// function doGet(e) {
//   const ss = SpreadsheetApp.openById(e.parameter.id);
//   const sh = ss.getSheetByName('旅途筆記') || ss.insertSheet('旅途筆記');
//   const data = sh.getRange(1,1).getValue();
//   return ContentService.createTextOutput(data || '{}').setMimeType(ContentService.MimeType.JSON);
// }
// function doPost(e) {
//   const p = JSON.parse(e.postData.contents);
//   const ss = SpreadsheetApp.openById(p.id);
//   const sh = ss.getSheetByName('旅途筆記') || ss.insertSheet('旅途筆記');
//   sh.getRange(1,1).setValue(p.data);
//   return ContentService.createTextOutput('ok');
// }
// ----------------------------------------

// 同步到 Google Sheet（多工作表版本）
// 每個工作表：A1 = JSON 備份、第3行起 = 人類可讀表格
async function syncToSheet() {
  if (!S.gasUrl || !S.sheetId) return;
  try {
    const payload = {
      id: S.sheetId,
      code: S.inviteCode || '',
      sheets: {
        '基本資訊': JSON.stringify(S.trip),
        '行程': JSON.stringify(S.itinerary),
        '航班': JSON.stringify(S.flights),
        '住宿': JSON.stringify(S.accommodation),
        '費用': JSON.stringify(S.expenses),
        '備忘錄': JSON.stringify(S.memos),
        '操作紀錄': JSON.stringify(S.history),
      },
      tables: {
        '行程': buildTable(['日期', '時間', '名稱', '地址', '類型', '停留(分)', '交通', '備註'], S.itinerary.map(i => [i.date, i.time, i.name, i.addr, { attraction: '景點', restaurant: '餐廳', shopping: '購物', other: '其他' }[i.type] || '其他', i.duration, i.transport || '', i.note])),
        '航班': buildTable(['航班號碼', '出發', '抵達', '起飛時間', '抵達時間', '備註'], S.flights.map(f => [f.number, f.from, f.to, f.depart, f.arrive, f.note])),
        '住宿': buildTable(['名稱', '地址', 'Check-in', 'Check-out', '備註'], S.accommodation.map(a => [a.name, a.addr, a.checkin, a.checkout, a.note])),
        '費用': buildTable(['日期', '說明', '類別', '金額', '幣別', '付款人'], S.expenses.map(e => [e.date, e.desc, e.category, e.amount, e.currency, e.payer])),
        '備忘錄': buildTable(['主題', '內容', '更新時間'], S.memos.map(m => [m.title, m.content, m.updatedAt])),
        '操作紀錄': buildTable(['時間', '操作'], S.history.map(h => [h.t, h.a])),
      },
    };
    await fetch(S.gasUrl, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify(payload),
    });
    S.lastSynced = nowStr();
    updateSyncStatus(`已同步至 Google Sheet ✓  ${S.lastSynced}`);
  } catch (e) {
    updateSyncStatus('⚠ 同步失敗（請確認 GAS URL 與 Sheet 權限）');
  }
}

function buildTable(headers, rows) {
  return [headers, ...rows];
}

async function syncFromSheet() {
  if (!S.gasUrl || !S.sheetId) return false;
  try {
    const resp = await fetch(`${S.gasUrl}?id=${S.sheetId}&action=read&code=${encodeURIComponent(S.inviteCode || '')}`);
    const text = await resp.text();
    if (!text || text === '{}') return false;
    const d = JSON.parse(text);
    // 新格式：每個 key 對應一個工作表的 JSON 字串
    if (d['基本資訊'] || d.trip) {
      const trip = d['基本資訊'] ? JSON.parse(d['基本資訊']) : d.trip;
      const itinerary = d['行程'] ? JSON.parse(d['行程']) : (d.itinerary || []);
      const flights = d['航班'] ? JSON.parse(d['航班']) : (d.flights || []);
      const accommodation = d['住宿'] ? JSON.parse(d['住宿']) : (d.accommodation || []);
      const expenses = d['費用'] ? JSON.parse(d['費用']) : (d.expenses || []);
      const history = d['操作紀錄'] ? JSON.parse(d['操作紀錄']) : (d.history || []);
      const memos = d['備忘錄'] ? JSON.parse(d['備忘錄']) : (d.memos || []);
      Object.assign(S, { trip, itinerary, flights, accommodation, expenses, history, memos });
      // ── Change 3: invalidate selected item if no longer exists ──
      if (S.selectedItemId && !S.itinerary.find(i => i.id === S.selectedItemId)) {
        S.selectedItemId = null;
      }
      saveToLocal();
      return true;
    }
  } catch (e) { console.warn('syncFromSheet error', e); }
  return false;
}

function updateSyncStatus(msg) {
  const el = $('sync-status');
  if (el) { el.textContent = msg; }
}

// 每次資料變動後，300ms 防抖後同步
function scheduleSave() {
  saveToLocal();
  if (!S.gasUrl || !S.sheetId) return;
  // 未通過白名單驗證：僅本機儲存，不同步至 GAS
  if (!S.isWhitelisted) {
    updateSyncStatus('⚠ 本機模式（Email 未在同步名單內）');
    return;
  }
  clearTimeout(S.syncTimer);
  S.syncTimer = setTimeout(() => {
    updateSyncStatus('同步中…');
    syncToSheet();
  }, 1200);
}

// ── URL QUERY 協作分享 ───────────────────
// 用 ?s= query string 取代 # hash，避免 LINE 截斷連結
function writeUrlHash() {
  if (S.sheetId) {
    const payload = {
      sid: S.sheetId,
      gas: S.gasUrl,
      ds: S.dataSource,
      th: S.trip.theme,
      st: S.trip.start,
      en: S.trip.end,
    };
    const encoded = btoa(unescape(encodeURIComponent(JSON.stringify(payload))))
      .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
    history.replaceState(null, '', '?s=' + encoded);
  } else {
    history.replaceState(null, '', window.location.pathname);
  }
}

function readUrlHash() {
  // 同時支援新格式 ?s= 與舊格式 #hash，向後相容
  const params = new URLSearchParams(window.location.search);
  const s = params.get('s') || window.location.hash.slice(1);
  if (!s) return null;
  try {
    const normalized = s.replace(/-/g, '+').replace(/_/g, '/');
    return JSON.parse(decodeURIComponent(escape(atob(normalized))));
  } catch (e) { return null; }
}

function generateShareUrl() {
  // 如果沒有設定 Sheet ID，就不給分享
  if (!S.sheetId) {
    toast('請先設定 Google Sheet 連結');
    return;
  }

  // 重新產生一份不含邀請碼的 payload，用 URL-safe base64 編碼
  // URL-safe base64 把 + 換成 -、/ 換成 _，LINE 等 app 才不會截斷連結
  const payload = {
    sid: S.sheetId,
    gas: S.gasUrl,
    ds: S.dataSource,
    th: S.trip.theme,
    st: S.trip.start,
    en: S.trip.end,
  };
  const encoded = btoa(unescape(encodeURIComponent(JSON.stringify(payload))))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

  // 用 ?s= 取代 # — LINE 遇到 # 會截斷，query string 不會
  let url = window.location.origin + window.location.pathname + '?s=' + encoded;

  // 6. 複製到剪貼簿並提醒使用者
  navigator.clipboard.writeText(url).then(() => {
    toast('✓ 分享連結已複製！（安全考量，連結不含邀請碼）');
  }).catch(() => {
    prompt('複製以下連結分享給同行成員（不含邀請碼）：', url);
  });
}

// ── INIT ──────────────────────────────────
async function initTrip(demo = false) {
  if (demo) {
    S.trip = { theme: '日本九州', start: '2025-03-15', end: '2025-03-22' };
    S.dataSource = '';
    S.sheetId = '';
    S.gasUrl = '';
    S.inviteCode = '';
    S.isWhitelisted = false;
    loadDemo();
  } else {
    const theme = v('trip-theme').trim();
    const start = v('trip-start');
    const end = v('trip-end');
    if (!theme || !start || !end) { toast('請填寫旅行主題和日期'); return; }
    S.trip = { theme, start, end };
    S.dataSource = v('data-source').trim();
    S.gasUrl = v('gas-url').trim();
    S.sheetId = extractSheetId(S.dataSource);
    const code = ($('invite-code-input') ? $('invite-code-input').value.trim() : '') ||
      localStorage.getItem('travel_invite_code') || '';
    S.inviteCode = code;

    // ── 邀請碼驗證 ──
    if (S.gasUrl && S.sheetId) {
      showInviteCodeStatus('驗證中…', 'info');
      const result = await verifyInviteCode(S.gasUrl, S.sheetId, code);
      if (result.reason === 'ok') {
        S.isWhitelisted = true;
        try { localStorage.setItem('travel_invite_code', code); } catch (e) { }
        try { localStorage.setItem('travel_gas_url', S.gasUrl); } catch (e) { }
        try { localStorage.setItem('travel_sheet_id', S.sheetId); } catch (e) { }
        showInviteCodeStatus('✅ 驗證通過，同步已啟用', 'ok');
      } else if (result.reason === 'wrong-code') {
        S.isWhitelisted = false;
        showInviteCodeStatus('❌ 邀請碼不正確，將以本機模式進入', 'error');
        await new Promise(r => setTimeout(r, 1200)); // 讓使用者看到提示
      } else if (result.reason === 'no-code') {
        S.isWhitelisted = false;
        showInviteCodeStatus('⚠ 未填邀請碼，將以本機模式進入', 'warn');
        await new Promise(r => setTimeout(r, 800));
      } else {
        // network-error：放行本機模式
        S.isWhitelisted = false;
        showInviteCodeStatus('⚠ 無法連線驗證，以本機模式進入', 'warn');
        await new Promise(r => setTimeout(r, 800));
      }
    } else {
      // 無 GAS：本機模式
      S.isWhitelisted = false;
    }

    // 從本機載入
    const local = loadFromLocal(S.sheetId);
    if (local && local.trip) {
      Object.assign(S, { itinerary: local.itinerary || [], flights: local.flights || [], accommodation: local.accommodation || [], expenses: local.expenses || [], memos: local.memos || [], history: local.history || [] });
      toast('已載入本機儲存的旅程資料');
    }

    // 從 Sheet 同步（驗證通過才嘗試）
    if (S.gasUrl && S.sheetId && S.isWhitelisted) {
      toast('正在從 Google Sheet 載入…');
      const ok = await syncFromSheet();
      if (ok) toast('已從 Google Sheet 同步行程 ✓');
    }

    // 寫入 URL hash 以便分享
    writeUrlHash();
  }

  closeModal('setup-modal');
  $('app').classList.remove('hidden');
  $('sidebar-theme').textContent = S.trip.theme;
  $('sidebar-dates').textContent = `${fmtShort(S.trip.start)} — ${fmtShort(S.trip.end)}`;
  // 同步手機 header
  if ($('mobile-theme')) $('mobile-theme').textContent = S.trip.theme;
  if ($('mobile-dates')) $('mobile-dates').textContent = `${fmtShort(S.trip.start)} — ${fmtShort(S.trip.end)}`;
  $('search-theme-hint').textContent = S.trip.theme;
  updateSyncUI();
  buildDateList();
  selectDate(getDefaultDate());
  renderFlights();
  renderAccom();
  renderExpenses();
  renderMemos();
  log('開始記錄旅程：' + S.trip.theme);
  startHeartbeat(); // 啟動協作者小活動

  // 自動從 Sheet 拉取最新資料（背景執行，每5分鐘更新）
  if (S.gasUrl && S.sheetId) {
    scheduleAutoRefresh();
  }
}

// 頁面載入時自動從 URL hash 填入所有設定欄位
window.addEventListener('DOMContentLoaded', () => {
  const hash = readUrlHash();
  if (!hash || !hash.sid) return;

  // ── 優先從 hash 帶入完整旅程設定（所有欄位）──
  const local = loadFromLocal(hash.sid);

  // 旅行主題：hash > localStorage
  const theme = hash.th || local?.trip?.theme || '';
  const start = hash.st || local?.trip?.start || '';
  const end = hash.en || local?.trip?.end || '';
  const ds = hash.ds || local?.dataSource || '';
  const gas = hash.gas || local?.gasUrl || '';

  $('trip-theme').value = theme;
  $('trip-start').value = start;
  $('trip-end').value = end;
  $('data-source').value = ds;
  $('gas-url').value = gas;

  // 顯示提示 banner
  const bar = $('share-auto-hint');
  const hint = $('share-auto-hint-text');
  if (bar && hint) {
    bar.style.display = 'flex';
    if (theme) {
      hint.textContent = `✓ 偵測到行程「${theme}」的分享連結，所有欄位已自動填入，確認後按「開始記錄旅程」`;
    } else {
      hint.textContent = '偵測到分享連結，Google Sheet 與 GAS 網址已填入，請補上旅行主題與日期後按開始';
    }
  }

  // 若有完整資訊（含 GAS），提供一鍵直接進入按鈕
  if (theme && start && end && gas) {
    const setupBox = document.querySelector('.setup-box');
    if (setupBox && !document.getElementById('auto-join-btn')) {
      const joinBtn = document.createElement('button');
      joinBtn.id = 'auto-join-btn';
      joinBtn.className = 'btn-primary';
      joinBtn.style.cssText = 'background:var(--teal);margin-bottom:8px;font-size:15px;';
      joinBtn.textContent = `⚡ 直接加入「${theme}」行程`;
      joinBtn.onclick = () => initTrip(false);
      // 插在第一個 btn-primary 之前
      const firstBtn = setupBox.querySelector('.btn-primary');
      if (firstBtn) setupBox.insertBefore(joinBtn, firstBtn);
    }
  }
});

function updateSyncUI() {
  const bar = $('sync-bar');
  if (!bar) return;
  if (S.sheetId) {
    bar.style.display = 'flex';
    $('sync-sheet-id').textContent = S.sheetId.slice(0, 10) + '…';
    if (!S.gasUrl) {
      updateSyncStatus('（未設定 GAS URL，僅本機儲存）');
    } else if (!S.isWhitelisted) {
      updateSyncStatus('⚠ 本機模式（Email 未在同步名單內）');
    } else {
      updateSyncStatus('等待變動後自動同步');
    }
    updateCollabLocal();
    updateAuthBadgeInApp();
  } else {
    bar.style.display = 'none';
  }
}

let _autoRefreshTimer = null;
// 每 5 分鐘自動從 Sheet 拉取最新資料（安靜後台更新）
function scheduleAutoRefresh() {
  if (_autoRefreshTimer) clearInterval(_autoRefreshTimer);
  _autoRefreshTimer = setInterval(async () => {
    if (!S.gasUrl || !S.sheetId) return;
    const ok = await syncFromSheet();
    if (ok) {
      const d = (S.selectedDate && S.selectedDate >= S.trip.start && S.selectedDate <= S.trip.end)
        ? S.selectedDate : getDefaultDate();
      buildDateList();
      selectDate(d);
      renderFlights(); renderAccom(); renderExpenses(); renderMemos();
      updateSyncStatus(`已從 Google Sheet 自動更新 ✓  ${nowStr()}`);
    }
  }, 5 * 60 * 1000); // 每 5 分鐘
}

function getDefaultDate() {
  const today = todayStr();
  if (today >= S.trip.start && today <= S.trip.end) return today;
  return S.trip.start;
}

// ── DATE LIST ─────────────────────────────
const WD = ['日', '一', '二', '三', '四', '五', '六'];

function buildDateList() {
  // Build sidebar list and mobile date dropdown
  const list = $('date-list');
  const dropdown = $('date-dropdown-mobile');
  list.innerHTML = '';
  if (dropdown) dropdown.innerHTML = '';
  const days = diffDays(S.trip.start, S.trip.end);
  for (let i = 0; i <= days; i++) {
    const d = addDays(S.trip.start, i);
    // Fix: parse date parts directly to avoid timezone offset shifting the day
    const [y, m, day] = d.split('-').map(Number);
    const wday = new Date(y, m - 1, day).getDay();
    // Sidebar button
    const btn = el('button', 'date-btn');
    btn.dataset.date = d;
    btn.innerHTML = `<span>Day${i + 1}　${d.slice(5).replace('-', '/')}</span>
                     <span class="day-sub">週${WD[wday]}</span>`;
    btn.onclick = () => selectDate(d);
    list.appendChild(btn);
    // Mobile dropdown option
    if (dropdown) {
      const opt = document.createElement('option');
      opt.value = d;
      opt.textContent = `Day${i + 1}  ${d.slice(5).replace('-', '/')} 週${WD[wday]}`;
      dropdown.appendChild(opt);
    }
  }
}

function selectDate(date) {
  S.selectedDate = date;
  document.querySelectorAll('.date-btn').forEach(b => b.classList.toggle('active', b.dataset.date === date));
  // Sync mobile dropdown
  const dropdown = $('date-dropdown-mobile');
  if (dropdown && dropdown.value !== date) dropdown.value = date;
  const [y, m, day] = date.split('-').map(Number);
  const wday = new Date(y, m - 1, day).getDay();
  const dayN = diffDays(S.trip.start, date) + 1;
  $('itinerary-date-title').textContent = `${date.slice(5).replace('-', '/')} 行程`;
  $('itinerary-day-label').textContent = `Day ${dayN}（週${WD[wday]}）`;
  setupMobileMap();
  renderItinerary();
}

// ── PANEL ─────────────────────────────────
function switchPanel(btn) {
  document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  btn.classList.add('active');
  $('panel-' + btn.dataset.panel).classList.add('active');
}

// ── ITINERARY RENDER ──────────────────────

// ── 全覽模式：所有日期行程 ──────────────────
function toggleAllDays() {
  S.showAllDays = !S.showAllDays;
  const btn = $('toggle-all-days-btn');
  if (btn) {
    btn.textContent = S.showAllDays ? '◈ 單日模式' : '▦ 全覽';
    btn.style.background = S.showAllDays ? 'var(--accent)' : '';
    btn.style.color = S.showAllDays ? '#fff' : '';
    btn.style.borderColor = S.showAllDays ? 'var(--accent)' : '';
  }
  const subtitle = $('itinerary-day-label');
  if (subtitle) subtitle.textContent = S.showAllDays ? `全程 ${diffDays(S.trip.start, S.trip.end) + 1} 天行程總覽` : '';
  renderItinerary();
}

function renderAllDays(list) {
  list.innerHTML = '';
  const typeLabel = { attraction: '景點', restaurant: '餐廳', shopping: '購物', other: '其他' };
  const days = diffDays(S.trip.start, S.trip.end);
  let hasAny = false;

  for (let i = 0; i <= days; i++) {
    const d = addDays(S.trip.start, i);
    const items = sortedItems(d);
    if (!items.length) continue;
    hasAny = true;

    const [y, m, day] = d.split('-').map(Number);
    const wday = new Date(y, m - 1, day).getDay();

    // Day header
    const hdr = document.createElement('div');
    hdr.className = 'all-days-hdr';
    hdr.innerHTML = `<span class="all-days-hdr-day">Day ${i + 1}</span>
      <span class="all-days-hdr-date">${d.slice(5).replace('-', '/')} 週${WD[wday]}</span>
      <span class="all-days-hdr-count">${items.length} 個行程</span>`;
    hdr.onclick = () => { S.showAllDays = false; selectDate(d); toggleAllDays(); };
    list.appendChild(hdr);

    items.forEach(item => {
      const wrap = document.createElement('div');
      wrap.className = 'itinerary-item' + (S.selectedItemId === item.id ? ' selected' : '');
      wrap.style.marginLeft = '8px';
      const expBadge = item.expenses?.length
        ? `<span class="item-exp-badge">💴 ${item.expenses.map(e => e.curr + e.amt).join('·')}</span>` : '';
      wrap.innerHTML = `
        <div class="item-time-col">
          <div class="item-time">${item.time || '--:--'}</div>
          <div class="item-type-dot dot-${item.type || 'other'}"></div>
        </div>
        <div class="item-right">
          <div class="item-body">
            <div class="item-name">${item.name}</div>
            <div class="item-meta">
              <span>${typeLabel[item.type] || '其他'}</span>
              ${item.duration ? `<span class="item-duration">⏱${item.duration}分</span>` : ''}
              ${expBadge}
            </div>
            ${item.note ? `<div class="item-note">${item.note}</div>` : ''}
            ${item.transport ? `<div class="item-transport">🚌 ${item.transport}</div>` : ''}
          </div>
          <div class="item-actions">
            <button class="item-action-btn" data-action="map" data-id="${item.id}">地圖</button>
            <button class="item-action-btn" data-action="edit" data-id="${item.id}">編輯</button>
            <button class="item-action-btn acc-btn" data-action="expense" data-id="${item.id}">記帳</button>
            <button class="item-action-btn del-btn" data-action="delete" data-id="${item.id}">刪除</button>
          </div>
        </div>`;
      wrap.addEventListener('click', e => {
        const btn = e.target.closest('[data-action]');
        if (btn) {
          const bid = btn.dataset.id;
          if (btn.dataset.action === 'map') { S.selectedDate = item.date; focusItem(bid); }
          if (btn.dataset.action === 'edit') openEditModal(bid);
          if (btn.dataset.action === 'expense') openExpenseForItem(bid);
          if (btn.dataset.action === 'delete') deleteItem(bid);
          return;
        }
        if (e.target.closest('.item-body')) { S.selectedDate = item.date; focusItem(item.id); }
      });
      list.appendChild(wrap);
    });
  }
  if (!hasAny) {
    list.innerHTML = '';
    const empty = document.createElement('div');
    empty.className = 'empty-state';
    empty.innerHTML = '<div class="empty-icon">◈</div><p>整個旅程尚無行程</p>';
    list.appendChild(empty);
  }
}

function renderItinerary() {
  const list = $('itinerary-list');

  // ── 全覽模式：顯示所有天的行程，依日期分組 ──
  if (S.showAllDays) {
    renderAllDays(list);
    return;
  }

  const items = sortedItems(S.selectedDate);
  if (!items.length) {
    list.innerHTML = '';
    list.appendChild($('itinerary-empty'));
    return;
  }
  list.innerHTML = '';
  const typeLabel = { attraction: '景點', restaurant: '餐廳', shopping: '購物', other: '其他' };

  items.forEach((item, idx) => {
    // Travel time connector
    if (idx > 0 && item.travelTime != null) {
      const mode = $('transport-mode').value;
      const modeIcon = { walking: '🚶', transit: '🚌', driving: '🚗' }[mode] || '🚶';
      const conn = el('div', 'travel-connector');
      conn.innerHTML = `<div class="tc-dot"></div>${modeIcon} 約 ${item.travelTime} 分鐘`;
      list.appendChild(conn);
    }

    const wrap = el('div', 'itinerary-item' + (S.selectedItemId === item.id ? ' selected' : ''));
    wrap.dataset.id = item.id;
    wrap.draggable = true;

    const expBadge = item.expenses?.length
      ? `<span class="item-exp-badge">💴 ${item.expenses.map(e => e.curr + e.amt).join('·')}</span>` : '';

    wrap.innerHTML = `
      <div class="drag-handle"></div>
      <div class="item-time-col">
        <div class="item-time">${item.time || '--:--'}</div>
        <div class="item-type-dot dot-${item.type || 'other'}"></div>
      </div>
      <div class="item-right">
        <div class="item-body">
          <div class="item-name">${item.name}</div>
          <div class="item-meta">
            <span>${typeLabel[item.type] || '其他'}</span>
            ${item.duration ? `<span class="item-duration">⏱${item.duration}分</span>` : ''}
            ${expBadge}
          </div>
          ${item.note ? `<div class="item-note">${item.note}</div>` : ''}
          ${item.transport ? `<div class="item-transport">🚌 ${item.transport}</div>` : ''}
        </div>
        <div class="item-actions">
          <button class="item-action-btn" data-action="map" data-id="${item.id}">地圖</button>
          <button class="item-action-btn" data-action="edit" data-id="${item.id}">編輯</button>
          <button class="item-action-btn acc-btn" data-action="expense" data-id="${item.id}">記帳</button>
          <button class="item-action-btn del-btn" data-action="delete" data-id="${item.id}">刪除</button>
        </div>
      </div>`;

    // 統一用 event delegation 操作按鈕
    wrap.addEventListener('click', e => {
      const btn = e.target.closest('[data-action]');
      if (btn) {
        const bid = btn.dataset.id;
        if (btn.dataset.action === 'map') focusItem(bid);
        if (btn.dataset.action === 'edit') openEditModal(bid);
        if (btn.dataset.action === 'expense') openExpenseForItem(bid);
        if (btn.dataset.action === 'delete') deleteItem(bid);
        return;
      }
      // 點擊行程名稱或 item-body 區域 → 地圖顯示
      if (e.target.closest('.item-body')) {
        focusItem(item.id);
        // 手機：自動滑動到地圖區域（在 focusItem 內處理）
      }
    });

    // Desktop Drag & Drop
    wrap.addEventListener('dragstart', e => { S.dragSrc = item.id; wrap.classList.add('dragging'); e.dataTransfer.effectAllowed = 'move'; });
    wrap.addEventListener('dragend', () => { wrap.classList.remove('dragging'); document.querySelectorAll('.itinerary-item').forEach(i => i.classList.remove('drag-over')); });
    wrap.addEventListener('dragover', e => { e.preventDefault(); wrap.classList.add('drag-over'); });
    wrap.addEventListener('dragleave', () => wrap.classList.remove('drag-over'));
    wrap.addEventListener('drop', e => { e.preventDefault(); wrap.classList.remove('drag-over'); reorder(S.dragSrc, item.id); });

    // Touch Drag & Drop（手機觸控）
    attachTouchDrag(wrap, item.id);

    list.appendChild(wrap);
  });
}

function sortedItems(date) {
  return S.itinerary.filter(i => i.date === date).sort((a, b) => (a.time || '').localeCompare(b.time || ''));
}

function reorder(srcId, tgtId) {
  const si = S.itinerary.findIndex(i => i.id === srcId);
  const ti = S.itinerary.findIndex(i => i.id === tgtId);
  if (si < 0 || ti < 0 || si === ti) return;
  const [m] = S.itinerary.splice(si, 1);
  S.itinerary.splice(ti, 0, m);
  renderItinerary();
  log('調整順序：' + m.name);
  scheduleSave();
}

// ── TOUCH DRAG & DROP（手機拖拉排序）────────
// 長按 400ms 啟動拖拉，移動時跟隨手指，放開後重新排序
function attachTouchDrag(el, itemId) {
  let longPressTimer = null;
  let ghost = null;
  let isDragging = false;
  let startY = 0;
  let offsetY = 0;

  el.addEventListener('touchstart', e => {
    // 如果觸碰在按鈕上，不啟動拖拉
    if (e.target.closest('button, [data-action]')) return;
    const touch = e.touches[0];
    startY = touch.clientY;

    longPressTimer = setTimeout(() => {
      isDragging = true;
      S.dragSrc = itemId;

      // 建立跟隨手指的幽靈元素
      ghost = el.cloneNode(true);
      const rect = el.getBoundingClientRect();
      offsetY = touch.clientY - rect.top;
      ghost.style.cssText = `
        position:fixed; left:${rect.left}px; top:${rect.top}px;
        width:${rect.width}px; opacity:0.85; z-index:9999;
        box-shadow:0 8px 32px rgba(0,0,0,0.22); border-radius:10px;
        pointer-events:none; transition:none;
        border:2px solid var(--accent);
      `;
      document.body.appendChild(ghost);
      el.classList.add('dragging');

      // 震動回饋（若裝置支援）
      if (navigator.vibrate) navigator.vibrate(40);
    }, 400);
  }, { passive: true });

  el.addEventListener('touchmove', e => {
    if (longPressTimer && !isDragging) {
      // 移動超過 8px 就取消長按
      if (Math.abs(e.touches[0].clientY - startY) > 8) {
        clearTimeout(longPressTimer);
        longPressTimer = null;
      }
      return;
    }
    if (!isDragging || !ghost) return;
    e.preventDefault();

    const touch = e.touches[0];
    const newTop = touch.clientY - offsetY;
    ghost.style.top = newTop + 'px';

    // 找出目前手指下方的行程 item
    ghost.style.display = 'none';
    const target = document.elementFromPoint(touch.clientX, touch.clientY);
    ghost.style.display = '';

    document.querySelectorAll('.itinerary-item').forEach(i => i.classList.remove('drag-over'));
    const overItem = target ? target.closest('.itinerary-item') : null;
    if (overItem && overItem !== el) {
      overItem.classList.add('drag-over');
      S.dragTarget = overItem.dataset.id;
    } else {
      S.dragTarget = null;
    }
  }, { passive: false });

  const endDrag = () => {
    clearTimeout(longPressTimer);
    longPressTimer = null;
    if (!isDragging) return;
    isDragging = false;

    if (ghost) { ghost.remove(); ghost = null; }
    el.classList.remove('dragging');
    document.querySelectorAll('.itinerary-item').forEach(i => i.classList.remove('drag-over'));

    if (S.dragTarget && S.dragTarget !== itemId) {
      reorder(itemId, S.dragTarget);
    }
    S.dragSrc = null;
    S.dragTarget = null;
  };

  el.addEventListener('touchend', endDrag, { passive: true });
  el.addEventListener('touchcancel', endDrag, { passive: true });
}

function deleteItem(id) {
  const item = S.itinerary.find(i => i.id === id);
  if (!item) return;
  showDeleteConfirm(item.name, () => {
    S.itinerary = S.itinerary.filter(i => i.id !== id);
    if (S.selectedItemId === id) { S.selectedItemId = null; resetMap(); }
    renderItinerary();
    toast('已刪除：' + item.name);
    log('刪除行程：' + item.name);
    scheduleSave();
  });
}

function showDeleteConfirm(name, onConfirm) {
  // Remove existing overlay if present
  const existing = document.getElementById('delete-confirm-overlay');
  if (existing) existing.remove();

  const overlay = document.createElement('div');
  overlay.id = 'delete-confirm-overlay';
  overlay.style.cssText = [
    'position:fixed', 'inset:0', 'background:rgba(0,0,0,.55)',
    'z-index:99999', 'display:flex', 'align-items:center', 'justify-content:center',
    'padding:20px', 'backdrop-filter:blur(2px)',
  ].join(';');

  overlay.innerHTML = `
    <div style="background:var(--paper);border-radius:16px;padding:28px 24px;max-width:340px;width:100%;
                box-shadow:0 24px 64px rgba(0,0,0,.3);animation:modalIn .2s ease;">
      <div style="font-size:16px;font-weight:700;font-family:'Playfair Display',serif;margin-bottom:8px;color:var(--ink)">確認刪除</div>
      <div style="font-size:13px;color:var(--ink-2);margin-bottom:22px;line-height:1.6">
        確定要刪除「<strong>${name}</strong>」？<br>此操作無法復原。
      </div>
      <div style="display:flex;gap:10px;justify-content:flex-end">
        <button id="del-cancel-btn" style="padding:10px 20px;border-radius:8px;border:1px solid var(--border);
                background:none;cursor:pointer;font-size:13px;color:var(--ink-2);font-family:'Noto Sans TC',sans-serif">取消</button>
        <button id="del-confirm-btn" style="padding:10px 20px;border-radius:8px;border:none;
                background:#c94f2a;color:#fff;cursor:pointer;font-size:13px;font-weight:600;
                font-family:'Noto Sans TC',sans-serif;min-width:72px">刪除</button>
      </div>
    </div>`;

  document.body.appendChild(overlay);

  const cleanup = () => { if (overlay.parentNode) overlay.remove(); };
  overlay.querySelector('#del-confirm-btn').addEventListener('click', () => { cleanup(); onConfirm(); });
  overlay.querySelector('#del-cancel-btn').addEventListener('click', cleanup);
  overlay.addEventListener('click', e => { if (e.target === overlay) cleanup(); });
  // Mobile: close on back gesture
  setTimeout(() => overlay.addEventListener('touchstart', e => { if (e.target === overlay) cleanup(); }, { passive: true }), 100);
}

// ── MOBILE MAP SETUP ─────────────────────
// On mobile, inject a compact map section below the itinerary list
function setupMobileMap() {
  // No-op: mobile map is now always-visible inline block, no dynamic injection needed
  return;
  const placeholder = $('mobile-map-placeholder');
  if (!placeholder || placeholder.dataset.built) return;
  placeholder.dataset.built = '1';
  placeholder.innerHTML = `
    <div class="mobile-inline-map-wrap" id="mobile-map-wrap" style="display:none">
      <div class="mobile-map-titlebar">
        <span id="mobile-map-location-label" class="mobile-map-loc-label">📍 點擊行程查看地圖</span>
        <div style="display:flex;gap:6px;align-items:center">
          <select id="transport-mode-mobile" class="transport-select" onchange="syncTransportModeMobile(this.value)">
            <option value="walking">🚶 步行</option>
            <option value="transit">🚌 大眾運輸</option>
            <option value="driving">🚗 乘車</option>
          </select>
          <button class="map-open-btn" onclick="openCurrentInGoogleMaps()" title="在 Google Maps 開啟">↗</button>
        </div>
      </div>
      <div class="mobile-map-frame-wrap">
        <div class="mobile-map-empty" id="mobile-map-empty-hint" style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;color:var(--ink-3);text-align:center;padding:20px">
          <div style="font-size:28px;opacity:.3;margin-bottom:10px">◈</div>
          <p style="font-size:12px;line-height:1.9">點擊上方行程名稱<br>地圖即連動顯示</p>
        </div>
        <iframe id="mobile-map-iframe" src="" allowfullscreen loading="lazy"
          referrerpolicy="no-referrer-when-downgrade"
          style="display:none;width:100%;height:100%;border:none">
        </iframe>
      </div>
      <div class="travel-time-bar" id="mobile-travel-time-bar" style="display:none">
        <span id="mobile-travel-time-icon"></span>
        <span id="mobile-travel-time-text"></span>
        <a id="mobile-travel-gmaps-link" href="#" target="_blank" class="travel-link">路線 →</a>
      </div>
    </div>`;
}

function syncTransportModeMobile(val) {
  const desktop = $('transport-mode');
  if (desktop && desktop.value !== val) desktop.value = val;
}

function setMobileMapIframe(url, label) {
  // No-op: replaced by inline-map-block approach
  return;
  // Works on all widths — desktop map-panel is shown by CSS, mobile placeholder is below list
  const wrap = $('mobile-map-wrap');
  const iframe = $('mobile-map-iframe');
  const hint = $('mobile-map-empty-hint');
  const lbl = $('mobile-map-location-label');
  if (!wrap || !iframe) return;
  wrap.style.display = 'block';
  iframe.src = url;
  iframe.style.display = 'block';
  if (hint) hint.style.display = 'none';
  if (lbl) lbl.textContent = label || '地圖';
}

function resetMobileMap() {
  const wrap = $('mobile-map-wrap');
  const iframe = $('mobile-map-iframe');
  const hint = $('mobile-map-empty-hint');
  const mttb = $('mobile-travel-time-bar');
  if (iframe) { iframe.style.display = 'none'; iframe.src = ''; }
  if (hint) hint.style.display = 'flex';
  if (mttb) mttb.style.display = 'none';
}

// ── MAP (iframe) ──────────────────────────
function buildMapUrl(query) {
  return `https://maps.google.com/maps?q=${encodeURIComponent(query)}&output=embed&hl=zh-TW&z=15`;
}

function buildMapUrlFromGmapLink(link) {
  const coordMatch = link.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (coordMatch) {
    const lat = coordMatch[1], lng = coordMatch[2];
    return `https://maps.google.com/maps?q=${lat},${lng}&output=embed&hl=zh-TW&z=15`;
  }
  const placeMatch = link.match(/place\/([^/@?]+)/);
  if (placeMatch) {
    return buildMapUrl(decodeURIComponent(placeMatch[1].replace(/\+/g, ' ')));
  }
  return null;
}

function showMapForQuery(query, label) {
  const url = buildMapUrl(query);
  setMapIframe(url, label);
}

function showMapForLink(link, label) {
  const url = buildMapUrlFromGmapLink(link) || buildMapUrl(label);
  setMapIframe(url, label);
}

function setMapIframe(url, label) {
  S.currentMapUrl = url;
  // ── 行程列表內嵌地圖（桌面 + 手機共用，CSS 統一控制高度）──
  const iframe = $('map-iframe');
  const hint = $('map-hint');
  if (iframe) {
    iframe.src = url;
    iframe.style.display = 'block';
    if (hint) hint.style.display = 'none';
  }
  const lbl = $('map-location-label');
  if (lbl) lbl.textContent = label || '地圖';
}

function resetMap() {
  const iframe = $('map-iframe');
  const hint = $('map-hint');
  if (iframe) { iframe.style.display = 'none'; iframe.src = ''; }
  if (hint) hint.style.display = 'flex';
  const ttb = $('travel-time-bar');
  if (ttb) ttb.style.display = 'none';
  const lbl = $('map-location-label');
  if (lbl) lbl.textContent = '點擊行程項目以查看位置';
  S.currentMapUrl = '';
}

function closeMobileMap() { /* legacy no-op */ }

function syncTransportMode(val) {
  const desktop = $('transport-mode');
  if (desktop && desktop.value !== val) desktop.value = val;
}

function focusItem(id) {
  S.selectedItemId = id;
  const items = sortedItems(S.selectedDate);
  const idx = items.findIndex(i => i.id === id);
  if (idx < 0) return;
  renderItinerary();

  const item = items[idx];
  const modeEl = $('transport-mode');
  const mode = modeEl ? modeEl.value : 'walking';
  const modeLabel = { walking: '步行', transit: '公共交通', driving: '乘車' }[mode];
  const modeIcon = { walking: '🚶', transit: '🚌', driving: '🚗' }[mode];

  let fromName = null, fromQ = null;

  if (idx === 0) {
    // 第一個行程：起始點為當天 check-in 的住宿（優先），其次為跨夜住宿
    // 若同一天有 checkout 和 check-in 重疊，以 check-in 當天那間為起點
    const checkinToday = S.accommodation.find(a =>
      a.checkin && a.checkin === S.selectedDate
    );
    const stayingToday = S.accommodation.find(a =>
      a.checkin && a.checkout &&
      a.checkin < S.selectedDate && a.checkout >= S.selectedDate
    );
    const accom = checkinToday || stayingToday;
    if (accom) {
      fromName = accom.name;
      fromQ = encodeURIComponent([accom.name, accom.addr, S.trip.theme].filter(Boolean).join(' '));
    }
  } else {
    // 第 n 個行程：從前一個行程出發
    const prev = items[idx - 1];
    fromName = prev.name;
    fromQ = encodeURIComponent([prev.name, prev.addr, S.trip.theme].filter(Boolean).join(' '));
  }

  const toQ = encodeURIComponent([item.name, item.addr, S.trip.theme].filter(Boolean).join(' '));

  if (fromQ) {
    // 顯示路線地圖
    const embedDir = `https://maps.google.com/maps?saddr=${fromQ}&daddr=${toQ}&output=embed&hl=zh-TW`;
    setMapIframe(embedDir, `${fromName} → ${item.name}`);

    const modeMap = { walking: 'walking', transit: 'transit', driving: 'driving' };
    const dirUrl = `https://www.google.com/maps/dir/?api=1&origin=${fromQ}&destination=${toQ}&travelmode=${modeMap[mode]}`;

    if ($('travel-time-icon')) $('travel-time-icon').textContent = modeIcon;
    if ($('travel-time-text')) $('travel-time-text').textContent = `${modeLabel}前往「${item.name}」`;
    if ($('travel-gmaps-link')) $('travel-gmaps-link').href = dirUrl;

    const ttb = $('travel-time-bar');
    if (ttb) ttb.style.display = 'flex';

    // (mobile travel bar merged into shared travel-time-bar)

    // 手機：滑動到地圖
    // Scroll to inline map block so user can see it
    const inlineMap = $('inline-map-itinerary');
    if (inlineMap) {
      setTimeout(() => inlineMap.scrollIntoView({ behavior: 'smooth', block: 'nearest' }), 150);
    }
  } else {
    // 無住宿資料：直接顯示景點位置
    if (item.gmaplink) showMapForLink(item.gmaplink, item.name);
    else showMapForQuery([item.name, item.addr, S.trip.theme].filter(Boolean).join(' '), item.name);

    const ttb = $('travel-time-bar');
    if (ttb) ttb.style.display = 'none';
    const mttb = $('mobile-travel-time-bar');
    if (mttb) mttb.style.display = 'none';
  }
}

function openCurrentInGoogleMaps() {
  const item = S.itinerary.find(i => i.id === S.selectedItemId);
  if (item) {
    const url = item.gmaplink || `https://www.google.com/maps/search/${encodeURIComponent([item.name, item.addr, S.trip.theme].filter(Boolean).join(' '))}`;
    window.open(url, '_blank');
  } else {
    window.open(`https://www.google.com/maps/search/${encodeURIComponent(S.trip.theme)}`, '_blank');
  }
}

// ── ADD SPOT（含 Modal 內地圖預覽）────────
function openAddSpotModal() {
  S.pendingSpot = null;
  $('spot-search-input').value = '';
  $('gmap-link-input').value = '';
  $('spot-search-results').classList.remove('visible');
  $('spot-search-results').innerHTML = '';
  $('spot-preview').style.display = 'none';
  $('spot-date').value = S.selectedDate || '';
  $('spot-time').value = '';
  $('spot-duration').value = '60';
  $('spot-note').value = '';
  $('spot-transport').value = '';
  $('spot-type').value = 'attraction';
  // 清空 modal 內地圖
  resetModalMap();
  switchSpotTab('search');
  openModal('modal-spot');
}

// Modal 內的地圖控制（與主地圖分開）
function setModalMapIframe(url, label) {
  const iframe = $('modal-map-iframe');
  const hint = $('modal-map-hint');
  const bar = $('modal-map-bar');
  if (!iframe) return;
  iframe.src = url;
  iframe.style.display = 'block';
  hint.style.display = 'none';
  bar.style.display = 'flex';
  $('modal-map-label').textContent = label || '地圖預覽';
}

function resetModalMap() {
  const iframe = $('modal-map-iframe');
  const hint = $('modal-map-hint');
  const bar = $('modal-map-bar');
  if (!iframe) return;
  iframe.src = '';
  iframe.style.display = 'none';
  hint.style.display = 'flex';
  if (bar) bar.style.display = 'none';
}

function openModalMapInGoogleMaps() {
  if (S.pendingSpot) {
    const url = S.pendingSpot.gmaplink ||
      `https://www.google.com/maps/search/${encodeURIComponent(S.pendingSpot.query || S.pendingSpot.name)}`;
    window.open(url, '_blank');
  }
}

function switchSpotTab(tab) {
  $('tab-search').classList.toggle('active', tab === 'search');
  $('tab-link').classList.toggle('active', tab === 'link');
  $('spot-tab-search').style.display = tab === 'search' ? '' : 'none';
  $('spot-tab-link').style.display = tab === 'link' ? '' : 'none';
}

function doSpotSearch() {
  const q = $('spot-search-input').value.trim();
  if (!q) return;
  const suggestions = generateSearchSuggestions(q);
  renderSpotSearchResults(suggestions);
}

function generateSearchSuggestions(q) {
  return [
    { name: q, addr: S.trip.theme + '（直接搜尋）', query: `${q} ${S.trip.theme}` },
    { name: `${q}（${S.trip.theme}）`, addr: '以旅行主題縮小範圍', query: `${q} ${S.trip.theme}` },
  ];
}

function renderSpotSearchResults(places) {
  const container = $('spot-search-results');
  container.innerHTML = '';
  places.forEach(p => {
    const item = document.createElement('div');
    item.className = 'search-result-item';
    item.innerHTML = `<div class="sr-name">${p.name}</div><div class="sr-addr">${p.addr || ''}</div>`;
    item.onclick = () => {
      S.pendingSpot = p;
      showSpotPreview(p);
      // ★ 修改 3：搜尋結果直接在 modal 內 iframe 顯示，不需跳轉
      const mapUrl = buildMapUrl(p.query || p.name);
      setModalMapIframe(mapUrl, p.name);
    };
    container.appendChild(item);
  });
  container.classList.add('visible');
}

function showSpotPreview(place) {
  const nameEl = $('preview-name');
  if (nameEl.tagName === 'INPUT') {
    nameEl.value = place.name || '';
    nameEl.placeholder = place.name ? '地點名稱（可修改）' : '請輸入景點名稱';
  } else {
    nameEl.textContent = place.name;
  }
  $('preview-addr').textContent = place.addr || '';
  $('spot-search-results').classList.remove('visible');
  $('spot-preview').style.display = 'block';
  // 若名稱為空，自動聚焦讓使用者填入
  if (!place.name && nameEl.tagName === 'INPUT') setTimeout(() => nameEl.focus(), 100);
}

// 從完整 Google Maps URL 解析地點名稱與座標
function extractGmapInfo(url) {
  let name = '', lat = '', lng = '';
  const placeMatch = url.match(/\/place\/([^/@?&]+)/);
  if (placeMatch) name = decodeURIComponent(placeMatch[1].replace(/\+/g, ' '));
  if (!name) { const q = url.match(/[?&]q=([^&]+)/); if (q) name = decodeURIComponent(q[1].replace(/\+/g, ' ')); }
  if (!name) { const s = url.match(/\/search\/([^/@?&]+)/); if (s) name = decodeURIComponent(s[1].replace(/\+/g, ' ')); }
  const coord = url.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (coord) { lat = coord[1]; lng = coord[2]; }
  return { name, lat, lng };
}

function buildEmbedFromInfo(info, fallbackLink) {
  if (info.lat && info.lng) return `https://maps.google.com/maps?q=${info.lat},${info.lng}&output=embed&hl=zh-TW&z=17`;
  if (info.name) return `https://maps.google.com/maps?q=${encodeURIComponent(info.name)}&output=embed&hl=zh-TW&z=16`;
  return null;
}

async function expandShortUrl(shortUrl) {
  // 用 allorigins proxy 取得 HTML，從 <title> 或 og:title 抓地點名稱
  // 同時 proxy 會 follow redirect，response 的 url 欄位就是展開後的長網址
  const proxyUrl = `https://api.allorigins.win/get?url=${encodeURIComponent(shortUrl)}`;
  const resp = await fetch(proxyUrl, { signal: AbortSignal.timeout(8000) });
  const data = await resp.json();
  // allorigins 回傳 { contents: '<html>...', status: { url: '展開後網址' } }
  const expandedUrl = data?.status?.url || '';
  const html = data?.contents || '';
  // 從 <title> 抓名稱，Google Maps title 格式通常是「地點名稱 - Google 地圖」
  const titleMatch = html.match(/<title[^>]*>([^<]+)<\/title>/i);
  let nameFromTitle = '';
  if (titleMatch) {
    nameFromTitle = titleMatch[1]
      .replace(/\s*[-–—]\s*Google.*/i, '')
      .replace(/Google 地圖|Google Maps/i, '')
      .trim();
  }
  return { expandedUrl, nameFromTitle };
}

async function parseGmapLink() {
  const link = $('gmap-link-input').value.trim();
  if (!link) { toast('請輸入連結'); return; }

  const isShort = /goo\.gl|maps\.app\.goo\.gl/i.test(link);

  // 先對本地 URL 試著解析
  let info = extractGmapInfo(link);
  let displayName = info.name;
  let finalLink = link;

  if (isShort || !displayName) {
    toast('📡 解析中…');
    try {
      const { expandedUrl, nameFromTitle } = await expandShortUrl(link);
      if (expandedUrl && expandedUrl !== link) {
        finalLink = expandedUrl;
        info = extractGmapInfo(expandedUrl);
      }
      // 名稱優先順序：URL 裡的 place/ > title > 空
      displayName = info.name || nameFromTitle || '';
    } catch (e) {
      console.warn('短網址展開失敗:', e);
    }
  }

  const embedUrl = buildEmbedFromInfo(info, finalLink);

  S.pendingSpot = {
    name: displayName || '',
    addr: (info.lat && info.lng) ? `${info.lat},${info.lng}` : '',
    gmaplink: finalLink,
    query: (displayName || '') + ' ' + S.trip.theme,
  };

  showSpotPreview(S.pendingSpot);

  if (embedUrl) {
    setModalMapIframe(embedUrl, displayName || '地圖預覽');
  }

  if (displayName) {
    toast('✓ 已解析：' + displayName);
  } else {
    const nameEl = $('preview-name');
    if (nameEl) setTimeout(() => nameEl.focus(), 100);
    toast('⚠ 無法自動取得名稱，請手動輸入');
  }
}

function addSpotToItinerary() {
  if (!S.pendingSpot) { toast('請先選擇地點'); return; }
  // 從 input 取最新名稱（使用者可能手動修改過）
  const nameEl = $('preview-name');
  const finalName = (nameEl && nameEl.tagName === 'INPUT' ? nameEl.value.trim() : '') || S.pendingSpot.name || '地圖地點';
  if (!finalName || finalName === '請輸入名稱') { toast('請輸入景點名稱'); if (nameEl) nameEl.focus(); return; }
  S.pendingSpot.name = finalName;
  const item = {
    id: uid(),
    name: finalName,
    addr: S.pendingSpot.addr || '',
    query: S.pendingSpot.query || S.pendingSpot.name + ' ' + S.trip.theme,
    gmaplink: S.pendingSpot.gmaplink || '',
    date: v('spot-date') || S.selectedDate,
    time: v('spot-time'),
    duration: parseInt(v('spot-duration')) || 60,
    type: v('spot-type'),
    note: v('spot-note'),
    transport: v('spot-transport') || '',
    expenses: [],
    travelTime: null,
  };
  S.itinerary.push(item);
  closeModal('modal-spot');
  if (item.date === S.selectedDate) renderItinerary();
  toast('已新增：' + item.name);
  log('新增行程：' + item.name + '（' + item.date + '）');
  scheduleSave();
}

// ── LIST IMPORT ───────────────────────────
function openGmapListModal() {
  $('gmaplist-date').value = S.selectedDate || '';
  $('gmaplist-text').value = '';
  openModal('modal-gmaplist');
}

// ★ 修改：移除無效的 URL 匯入（Google Maps 清單因 CORS 無法抓取）
// 僅保留文字清單匯入，每行一個地點名稱
function importList() {
  const date = v('gmaplist-date') || S.selectedDate;
  if (!date) { toast('請選擇要加入的日期'); return; }
  const textVal = v('gmaplist-text').trim();
  if (!textVal) { toast('請輸入地點清單（每行一個地點）'); return; }
  const names = textVal.split('\n').map(s => s.trim()).filter(Boolean);
  if (!names.length) { toast('清單內容為空'); return; }
  names.forEach(name => {
    S.itinerary.push({
      id: uid(), name, addr: '', query: name + ' ' + S.trip.theme, gmaplink: '',
      date, time: '', duration: 60, type: 'attraction', note: '從清單匯入', expenses: [], travelTime: null,
    });
  });
  closeModal('modal-gmaplist');
  selectDate(date);
  toast(`✓ 已匯入 ${names.length} 個地點`);
  log(`批次匯入 ${names.length} 個地點`);
  scheduleSave();
}

// ── FLIGHTS ───────────────────────────────
function renderFlights() {
  const el = $('flights-list');
  if (!S.flights.length) { el.innerHTML = '<div class="empty-state"><div class="empty-icon">◎</div><p>尚未新增航班</p></div>'; return; }
  el.innerHTML = S.flights.map(f => `
    <div class="flight-card">
      <div class="flight-airport">
        <div class="flight-code">${f.from}</div>
        <div class="flight-time">${fmtDT(f.depart)}</div>
      </div>
      <div class="flight-arrow">✈<div class="flight-num">${f.number}</div>${f.note ? `<div style="font-size:10px;color:var(--ink-3);margin-top:3px">${f.note}</div>` : ''}</div>
      <div class="flight-airport">
        <div class="flight-code">${f.to}</div>
        <div class="flight-time">${fmtDT(f.arrive)}</div>
      </div>
      <div style="display:flex;flex-direction:column;gap:6px">
        <button class="btn-sm btn-outline" onclick="editFlight('${f.id}')">編輯</button>
        <button class="btn-sm btn-danger" onclick="delFlight('${f.id}')">刪除</button>
      </div>
    </div>`).join('');
}

function openAddFlightModal() {
  $('flight-edit-id').value = '';
  ['flight-number', 'flight-from', 'flight-to', 'flight-depart', 'flight-arrive', 'flight-note'].forEach(id => $(id).value = '');
  $('flight-modal-title').textContent = '新增航班';
  openModal('modal-flight');
}

function editFlight(id) {
  const f = S.flights.find(x => x.id === id);
  if (!f) return;
  $('flight-edit-id').value = id;
  $('flight-number').value = f.number || '';
  $('flight-from').value = f.from || '';
  $('flight-to').value = f.to || '';
  $('flight-depart').value = f.depart || '';
  $('flight-arrive').value = f.arrive || '';
  $('flight-note').value = f.note || '';
  $('flight-modal-title').textContent = '編輯航班';
  openModal('modal-flight');
}

function saveFlightModal() {
  const editId = v('flight-edit-id');
  const data = {
    number: v('flight-number'), from: v('flight-from'), to: v('flight-to'),
    depart: v('flight-depart'), arrive: v('flight-arrive'), note: v('flight-note'),
  };
  if (!data.number || !data.from || !data.to) { toast('請填寫航班資訊'); return; }
  if (editId) {
    const f = S.flights.find(x => x.id === editId);
    if (f) Object.assign(f, data);
    toast('航班已更新：' + data.number);
    log('更新航班：' + data.number);
  } else {
    S.flights.push({ id: uid(), ...data });
    toast('航班已儲存');
    log('新增航班：' + data.number);
  }
  renderFlights();
  closeModal('modal-flight');
  ['flight-number', 'flight-from', 'flight-to', 'flight-depart', 'flight-arrive', 'flight-note'].forEach(id => $(id).value = '');
  scheduleSave();
}

function delFlight(id) { S.flights = S.flights.filter(f => f.id !== id); renderFlights(); toast('航班已刪除'); scheduleSave(); }

// ── ACCOMMODATION ─────────────────────────
function openAddAccomModal() {
  $('accom-edit-id').value = '';
  ['accom-name-input', 'accom-addr-input', 'accom-note', 'accom-gmaplink', 'accom-checkin', 'accom-checkout']
    .forEach(id => $(id).value = '');
  $('accom-modal-title').textContent = '新增住宿';
  openModal('modal-accom');
}

function editAccom(id) {
  const a = S.accommodation.find(x => x.id === id);
  if (!a) return;
  $('accom-edit-id').value = id;
  $('accom-name-input').value = a.name || '';
  $('accom-addr-input').value = a.addr || '';
  $('accom-checkin').value = a.checkin || '';
  $('accom-checkout').value = a.checkout || '';
  $('accom-note').value = a.note || '';
  $('accom-gmaplink').value = a.gmaplink || '';
  $('accom-modal-title').textContent = '編輯住宿';
  openModal('modal-accom');
}

function saveAccomModal() {
  const name = v('accom-name-input');
  if (!name) { toast('請填寫住宿名稱'); return; }
  const editId = v('accom-edit-id');
  const data = {
    name, addr: v('accom-addr-input'),
    checkin: v('accom-checkin'), checkout: v('accom-checkout'),
    note: v('accom-note'), gmaplink: v('accom-gmaplink'),
  };
  if (editId) {
    const a = S.accommodation.find(x => x.id === editId);
    if (a) Object.assign(a, data);
    toast('住宿已更新：' + name);
    log('更新住宿：' + name);
  } else {
    S.accommodation.push({ id: uid(), ...data });
    toast('住宿已儲存：' + name);
    log('新增住宿：' + name);
  }
  renderAccom();
  closeModal('modal-accom');
  scheduleSave();
}

function renderAccom() {
  const el = $('accom-list');
  if (!S.accommodation.length) { el.innerHTML = '<div class="empty-state"><div class="empty-icon">◻</div><p>尚未新增住宿</p></div>'; return; }
  el.innerHTML = S.accommodation.map(a => {
    const mapUrl = a.gmaplink || `https://www.google.com/maps/search/${encodeURIComponent((a.name + ' ' + a.addr + ' ' + S.trip.theme).trim())}`;
    return `<div class="accom-card">
      <div class="accom-icon">🏨</div>
      <div class="accom-info">
        <div class="accom-name">${a.name}</div>
        <div class="accom-dates">Check-in: ${a.checkin || '—'}　Check-out: ${a.checkout || '—'}</div>
        ${a.addr ? `<div class="accom-note">📍 ${a.addr}</div>` : ''}
        ${a.note ? `<div class="accom-note">${a.note}</div>` : ''}
        <a class="accom-map-link" href="${mapUrl}" target="_blank" onclick="showAccomMap('${a.id}');return false;">在地圖查看 →</a>
      </div>
      <div class="accom-actions">
        <button class="btn-sm btn-outline" onclick="editAccom('${a.id}')">編輯</button>
        <button class="btn-sm btn-danger" onclick="delAccom('${a.id}')">刪除</button>
      </div>
    </div>`;
  }).join('');
}


// ── 住宿頁面內嵌地圖控制 ──────────────────────────────
function showAccomInlineMap(url, label) {
  const iframe = $('accom-map-iframe');
  const hint = $('accom-map-hint');
  if (!iframe) return;
  iframe.src = url;
  iframe.style.display = 'block';
  if (hint) hint.style.display = 'none';
  const lbl = $('accom-map-label');
  if (lbl) lbl.textContent = label || '地圖';
}

function openCurrentAccomInGoogleMaps() {
  if (S._lastAccomMapUrl) {
    window.open(S._lastAccomMapUrl.replace('&output=embed', ''), '_blank');
  } else {
    window.open('https://www.google.com/maps/search/' + encodeURIComponent(S.trip.theme + ' 住宿'), '_blank');
  }
}

function showAccomMap(id) {
  const a = S.accommodation.find(x => x.id === id);
  if (!a) return;
  const q = [a.name, a.addr, S.trip.theme].filter(Boolean).join(' ');
  const url = a.gmaplink ? (buildMapUrlFromGmapLink(a.gmaplink) || buildMapUrl(q)) : buildMapUrl(q);
  S._lastAccomMapUrl = url;
  showAccomInlineMap(url, a.name);
}

function delAccom(id) { S.accommodation = S.accommodation.filter(x => x.id !== id); renderAccom(); toast('住宿已刪除'); scheduleSave(); }

// ── EXPENSES ──────────────────────────────
function openAddExpenseModal(linkedId = '') {
  $('expense-modal-title').textContent = '新增費用';
  $('expense-edit-id').value = '';
  $('exp-date').value = S.selectedDate || '';
  $('exp-category').value = 'food';
  $('exp-desc').value = '';
  $('exp-amount').value = '';
  $('exp-currency').value = 'JPY';
  $('exp-payer').value = '';
  $('exp-splits').innerHTML = '';
  $('exp-linked-item').value = linkedId;
  if (linkedId) {
    const item = S.itinerary.find(i => i.id === linkedId);
    if (item) $('exp-desc').value = item.name;
  }
  openModal('modal-expense');
}

function editExpense(id) {
  const e = S.expenses.find(x => x.id === id);
  if (!e) return;
  $('expense-modal-title').textContent = '編輯費用';
  $('expense-edit-id').value = id;
  $('exp-date').value = e.date || '';
  $('exp-category').value = e.category || 'food';
  $('exp-desc').value = e.desc || '';
  $('exp-amount').value = e.amount || '';
  $('exp-currency').value = e.currency || 'JPY';
  $('exp-payer').value = e.payer || '';
  $('exp-linked-item').value = e.linkedItemId || '';
  $('exp-splits').innerHTML = '';
  (e.splits || []).forEach(s => {
    const row = document.createElement('div');
    row.className = 'split-row';
    row.innerHTML = `<input type="text" placeholder="姓名" value="${s.name}" /><input type="number" placeholder="金額" min="0" value="${s.amount}" /><button class="split-remove" onclick="this.parentElement.remove()">✕</button>`;
    $('exp-splits').appendChild(row);
  });
  openModal('modal-expense');
}

function openExpenseForItem(id) { openAddExpenseModal(id); }

function addSplitRow() {
  const row = document.createElement('div');
  row.className = 'split-row';
  row.innerHTML = `<input type="text" placeholder="姓名" /><input type="number" placeholder="金額" min="0" /><button class="split-remove" onclick="this.parentElement.remove()">✕</button>`;
  $('exp-splits').appendChild(row);
}

function saveExpenseModal() {
  const amount = parseFloat(v('exp-amount'));
  if (!amount) { toast('請輸入金額'); return; }
  const splits = [...document.querySelectorAll('#exp-splits .split-row')].map(r => {
    const ins = r.querySelectorAll('input');
    return { name: ins[0].value, amount: parseFloat(ins[1].value) || 0 };
  }).filter(s => s.name);
  const linkedId = v('exp-linked-item');
  const editId = v('expense-edit-id');
  const data = {
    date: v('exp-date'), category: v('exp-category'), desc: v('exp-desc'),
    amount, currency: v('exp-currency'), payer: v('exp-payer'), splits, linkedItemId: linkedId,
  };
  if (editId) {
    const e = S.expenses.find(x => x.id === editId);
    if (e) Object.assign(e, data);
    toast('費用已更新');
    log(`更新費用：${data.desc} ${data.amount}${data.currency}`);
  } else {
    const exp = { id: uid(), ...data };
    S.expenses.push(exp);
    if (linkedId) {
      const item = S.itinerary.find(i => i.id === linkedId);
      if (item) { if (!item.expenses) item.expenses = []; item.expenses.push({ curr: exp.currency, amt: exp.amount }); renderItinerary(); }
    }
    toast('費用已記錄');
    log(`記錄費用：${data.desc} ${data.amount}${data.currency}`);
  }
  renderExpenses();
  closeModal('modal-expense');
  scheduleSave();
}

function renderExpenses() {
  const catIcons = { food: '🍽', transport: '🚌', attraction: '🎫', accommodation: '🏨', shopping: '🛍', other: '💰' };
  const catNames = { food: '餐飲', transport: '交通', attraction: '景點', accommodation: '住宿', shopping: '購物', other: '其他' };

  // ── 多幣別各自加總 ──
  const byCurr = {};
  S.expenses.forEach(e => { byCurr[e.currency] = (byCurr[e.currency] || 0) + e.amount; });
  const currSummary = Object.entries(byCurr).map(([c, v]) => `${c} ${v.toLocaleString()}`).join('　');
  $('total-amount').textContent = currSummary || '0';
  $('total-count').textContent = S.expenses.length;

  // ── 各付款人應付小計（更新摘要區） ──
  const payerTotals = {};
  S.expenses.forEach(e => {
    if (e.payer) payerTotals[e.payer] = (payerTotals[e.payer] || {});
    if (e.payer) payerTotals[e.payer][e.currency] = (payerTotals[e.payer][e.currency] || 0) + e.amount;
  });
  let payerEl = $('payer-summary');
  if (!payerEl) {
    payerEl = document.createElement('div');
    payerEl.id = 'payer-summary';
    payerEl.className = 'payer-summary';
    const summaryGrid = document.querySelector('.expense-summary');
    if (summaryGrid) summaryGrid.insertAdjacentElement('afterend', payerEl);
  }
  if (Object.keys(payerTotals).length) {
    payerEl.innerHTML = '<div class="payer-summary-title">各人已付小計</div>' +
      Object.entries(payerTotals).map(([name, currs]) =>
        `<div class="payer-row"><span class="payer-name">${name}</span><span class="payer-amounts">${Object.entries(currs).map(([c, v]) => `${c} ${v.toLocaleString()}`).join('　')
        }</span></div>`
      ).join('');
    payerEl.style.display = 'block';
  } else {
    payerEl.style.display = 'none';
  }

  const el = $('expense-list');
  if (!S.expenses.length) { el.innerHTML = '<div class="empty-state"><div class="empty-icon">◇</div><p>尚未記錄費用</p></div>'; return; }
  el.innerHTML = S.expenses.map(e => `
    <div class="expense-item">
      <div class="exp-cat">${catIcons[e.category] || '💰'}</div>
      <div class="exp-info">
        <div class="exp-desc-text">${e.desc}${e.linkedItemId ? '<span class="linked-badge">行程</span>' : ''}</div>
        <div class="exp-meta">${e.date}·${catNames[e.category]}${e.payer ? '·' + e.payer + '付' : ''}</div>
        ${(e.splits || []).length ? `<div class="exp-meta">${e.splits.map(s => s.name + ' ' + s.amount).join('/')}</div>` : ''}
      </div>
      <div class="exp-amount-col">
        <div class="exp-amount-val">${e.currency} ${e.amount.toLocaleString()}</div>
      </div>
      <div style="display:flex;gap:4px">
        <button class="btn-sm btn-outline" onclick="editExpense('${e.id}')">編輯</button>
        <button class="btn-sm btn-danger" onclick="delExpense('${e.id}')">✕</button>
      </div>
    </div>`).join('');
}
function delExpense(id) { S.expenses = S.expenses.filter(e => e.id !== id); renderExpenses(); toast('已刪除費用'); scheduleSave(); }

// ── EDIT ──────────────────────────────────
function openEditModal(id) {
  const item = S.itinerary.find(i => i.id === id);
  if (!item) return;
  $('edit-item-id').value = id;
  $('edit-name').value = item.name;
  $('edit-addr').value = item.addr || '';
  $('edit-gmaplink').value = item.gmaplink || '';
  $('edit-date').value = item.date;
  $('edit-time').value = item.time || '';
  $('edit-duration').value = item.duration || '';
  $('edit-note').value = item.note || '';
  $('edit-transport').value = item.transport || '';
  openModal('modal-edit');
}

function saveEdit() {
  const id = v('edit-item-id');
  const item = S.itinerary.find(i => i.id === id);
  if (!item) return;
  item.name = v('edit-name');
  item.addr = v('edit-addr');
  item.gmaplink = v('edit-gmaplink');
  item.query = [item.name, item.addr, S.trip.theme].filter(Boolean).join(' ');
  item.date = v('edit-date');
  item.time = v('edit-time');
  item.duration = parseInt(v('edit-duration')) || 0;
  item.note = v('edit-note');
  item.transport = v('edit-transport') || '';
  closeModal('modal-edit');
  renderItinerary();
  toast('已更新行程');
  log('編輯行程：' + item.name);
  scheduleSave();
}

// ── EXPORT ────────────────────────────────
function exportToMarkdown() {
  let md = `# ${S.trip.theme} 旅遊紀錄\n**旅程日期**：${S.trip.start} ～ ${S.trip.end}\n\n`;
  if (S.flights.length) {
    md += `## ✈ 航班\n`;
    S.flights.forEach(f => { md += `- **${f.number}** ${f.from}→${f.to}　起飛：${fmtDT(f.depart)}　抵達：${fmtDT(f.arrive)}${f.note ? '　' + f.note : ''}\n`; });
    md += '\n';
  }
  if (S.accommodation.length) {
    md += `## 🏨 住宿\n`;
    S.accommodation.forEach(a => { md += `- **${a.name}**　${a.checkin}～${a.checkout}${a.note ? '　' + a.note : ''}${a.gmaplink ? '\n  ' + a.gmaplink : ''}\n`; });
    md += '\n';
  }
  const days = diffDays(S.trip.start, S.trip.end);
  for (let i = 0; i <= days; i++) {
    const d = addDays(S.trip.start, i);
    const items = sortedItems(d);
    if (!items.length) continue;
    md += `## Day ${i + 1}（${d}）\n`;
    const tl = { attraction: '景點', restaurant: '餐廳', shopping: '購物', other: '其他' };
    items.forEach(item => {
      md += `### ${item.time || '--:--'} ${item.name}`;
      if (item.type) md += ` · ${tl[item.type] || ''}`;
      md += '\n';
      if (item.addr) md += `📍 ${item.addr}\n`;
      if (item.duration) md += `⏱ 停留 ${item.duration} 分鐘\n`;
      if (item.note) md += `> ${item.note}\n`;
      if (item.gmaplink) md += `🗺 ${item.gmaplink}\n`;
      const linked = S.expenses.filter(e => e.linkedItemId === item.id);
      if (linked.length) md += `💴 ${linked.map(e => e.desc + ' ' + e.currency + e.amount.toLocaleString()).join('、')}\n`;
      md += '\n';
    });
  }
  if (S.expenses.length) {
    const byCurr = {};
    S.expenses.forEach(e => { byCurr[e.currency] = (byCurr[e.currency] || 0) + e.amount; });
    const totalStr = Object.entries(byCurr).map(([c, v]) => `${c} ${v.toLocaleString()}`).join('、');
    md += `## 💰 費用明細（總計：${totalStr}）\n\n`;
    md += `| 日期 | 說明 | 類別 | 金額 | 幣別 | 付款人 |\n|------|------|------|------|------|-----------|\n`;
    const catN2 = { food: '餐飲', transport: '交通', attraction: '景點', accommodation: '住宿', shopping: '購物', other: '其他' };
    S.expenses.forEach(e => { md += `| ${e.date} | ${e.desc} | ${catN2[e.category] || '其他'} | ${e.amount.toLocaleString()} | ${e.currency} | ${e.payer || '—'} |\n`; });
    const payerMap2 = {};
    S.expenses.forEach(e => { if (e.payer) { payerMap2[e.payer] = payerMap2[e.payer] || {}; payerMap2[e.payer][e.currency] = (payerMap2[e.payer][e.currency] || 0) + e.amount; } });
    if (Object.keys(payerMap2).length) {
      md += `\n**各人已付：** ${Object.entries(payerMap2).map(([n, cs]) => n + '：' + Object.entries(cs).map(([c, v]) => c + ' ' + v.toLocaleString()).join('/')).join('　')}\n`;
    }
    md += '\n';
  }
  md += `## 📋 操作紀錄\n`;
  S.history.slice(0, 30).forEach(h => { md += `- ${h.t}　${h.a}\n`; });
  blob(md, 'text/markdown', S.trip.theme + '_旅程紀錄.md');
  toast('Markdown 已下載');
}

// 匯出行程為可列印的乾淨版本
function exportToPDF() {
  // 動態產生列印專用的行程摘要頁
  const days = diffDays(S.trip.start, S.trip.end);
  const tl = { attraction: '景點', restaurant: '餐廳', shopping: '購物', other: '其他' };
  let html = `<!DOCTYPE html><html lang="zh-TW"><head><meta charset="UTF-8">
<title>${S.trip.theme} 旅程行程表</title>
<style>
  body{font-family:'Noto Sans TC',sans-serif;max-width:800px;margin:0 auto;padding:30px;color:#1a1614;}
  h1{font-size:24px;border-bottom:2px solid #c94f2a;padding-bottom:10px;color:#c94f2a;}
  h2{font-size:16px;background:#f8f4ef;padding:8px 14px;border-radius:6px;margin:20px 0 10px;}
  h3{font-size:14px;color:#4a3f3a;margin:4px 0;}
  .meta{font-size:12px;color:#8a7b75;margin-bottom:6px;}
  .item{padding:10px 14px;border-left:3px solid #e6ddd2;margin-bottom:8px;}
  .item.restaurant{border-color:#b8913a;} .item.attraction{border-color:#c94f2a;}
  .note{font-size:11px;color:#8a7b75;font-style:italic;margin-top:4px;}
  .flight-row,.accom-row,.exp-row{font-size:13px;padding:6px 0;border-bottom:1px solid #e6ddd2;}
  table{width:100%;border-collapse:collapse;font-size:12px;margin-top:8px;}
  th{background:#1a1614;color:#fff;padding:6px 10px;text-align:left;}
  td{padding:5px 10px;border-bottom:1px solid #e6ddd2;}
  @media print{@page{margin:15mm}}
</style></head><body>`;
  html += `<h1>✦ ${S.trip.theme}</h1><div class="meta">📅 ${S.trip.start} ～ ${S.trip.end}</div>`;
  if (S.flights.length) {
    html += `<h2>✈ 航班資訊</h2>`;
    S.flights.forEach(f => { html += `<div class="flight-row"><b>${f.number}</b> ${f.from} → ${f.to}　起飛 ${fmtDT(f.depart)}　抵達 ${fmtDT(f.arrive)}${f.note ? ' (' + f.note + ')' : ''}</div>`; });
  }
  if (S.accommodation.length) {
    html += `<h2>🏨 住宿資訊</h2>`;
    S.accommodation.forEach(a => { html += `<div class="accom-row"><b>${a.name}</b>　${a.checkin} ～ ${a.checkout}${a.note ? ' ・ ' + a.note : ''}</div>`; });
  }
  for (let i = 0; i <= days; i++) {
    const d = addDays(S.trip.start, i);
    const items = sortedItems(d);
    if (!items.length) continue;
    const [yd, md2, dd] = d.split('-').map(Number);
    const wday = new Date(yd, md2 - 1, dd).getDay();
    html += `<h2>Day ${i + 1}（${d} 週${WD[wday]}）</h2>`;
    items.forEach(item => {
      html += `<div class="item ${item.type || 'other'}"><h3>${item.time || '--:--'} ${item.name} <span style="font-weight:400;font-size:12px;color:#8a7b75">・${tl[item.type] || '其他'}${item.duration ? ' ⏱' + item.duration + '分' : ''}</span></h3>`;
      if (item.addr) html += `<div class="meta">📍 ${item.addr}</div>`;
      if (item.note) html += `<div class="note">${item.note}</div>`;
      html += `</div>`;
    });
  }
  // ── 費用明細 ──
  if (S.expenses.length) {
    const byCurr = {};
    S.expenses.forEach(e => { byCurr[e.currency] = (byCurr[e.currency] || 0) + e.amount; });
    const totalStr = Object.entries(byCurr).map(([c, v]) => `${c} ${v.toLocaleString()}`).join('　');
    html += `<h2>💰 費用明細（總計：${totalStr}）</h2>
    <table><tr><th>日期</th><th>說明</th><th>類別</th><th>金額</th><th>幣別</th><th>付款人</th><th>分帳</th></tr>`;
    S.expenses.forEach(e => {
      const catN = { food: '餐飲', transport: '交通', attraction: '景點', accommodation: '住宿', shopping: '購物', other: '其他' }[e.category] || '其他';
      const splits = (e.splits || []).length ? e.splits.map(s => s.name + ' ' + s.amount).join('/') : '';
      html += `<tr><td>${e.date}</td><td>${e.desc}</td><td>${catN}</td><td>${e.amount.toLocaleString()}</td><td>${e.currency}</td><td>${e.payer || '—'}</td><td>${splits}</td></tr>`;
    });
    // per-payer totals
    const payerMap = {};
    S.expenses.forEach(e => { if (e.payer) { payerMap[e.payer] = payerMap[e.payer] || {}; payerMap[e.payer][e.currency] = (payerMap[e.payer][e.currency] || 0) + e.amount; } });
    if (Object.keys(payerMap).length) {
      html += `</table><div style="margin-top:10px;font-size:13px;color:#4a3f3a"><b>各人已付：</b> `;
      html += Object.entries(payerMap).map(([n, cs]) => `${n}：${Object.entries(cs).map(([c, v]) => c + ' ' + v.toLocaleString()).join('/')}`).join('　');
      html += `</div>`;
    } else {
      html += `</table>`;
    }
  }

  html += `</body></html>`;
  const win = window.open('', '_blank');
  win.document.write(html);
  win.document.close();
  setTimeout(() => win.print(), 500);
  toast('行程表已在新視窗開啟，可列印或存 PDF');
}

function exportToJSON() {
  const data = JSON.stringify({ trip: S.trip, itinerary: S.itinerary, flights: S.flights, accommodation: S.accommodation, expenses: S.expenses, history: S.history }, null, 2);
  blob(data, 'application/json', S.trip.theme + '_備份.json');
  toast('備份檔案已下載');
  log('匯出備份');
}

function importJSON(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const d = JSON.parse(ev.target.result);
      Object.assign(S, d);
      S.sheetId = extractSheetId(S.dataSource || '');
      S.memos = S.memos || [];
      closeModal('setup-modal');
      $('app').classList.remove('hidden');
      $('sidebar-theme').textContent = S.trip.theme;
      $('sidebar-dates').textContent = `${fmtShort(S.trip.start)} — ${fmtShort(S.trip.end)}`;
      $('search-theme-hint').textContent = S.trip.theme;
      updateSyncUI();
      buildDateList();
      selectDate(S.selectedDate || getDefaultDate());
      renderFlights(); renderAccom(); renderExpenses(); renderMemos();
      toast('備份已還原：' + S.trip.theme);
      scheduleSave();
    } catch { toast('檔案格式有誤'); }
  };
  reader.readAsText(file);
}

// ── 手動觸發 Sheet 同步按鈕 ──────────────
async function manualSync() {
  if (!S.gasUrl || !S.sheetId) {
    toast('請先設定 Google Sheet 網址與 GAS URL');
    return;
  }
  if (!S.isWhitelisted) {
    toast('⚠ 您的 Email 不在同步名單內，無法上傳至 Google Sheet');
    return;
  }
  updateSyncStatus('同步中…');
  await syncToSheet();
}

// ── 從 Sheet 拉取最新資料 ────────────────
async function pullFromSheet() {
  if (!S.gasUrl || !S.sheetId) { toast('尚未設定 GAS URL，無法拉取'); return; }
  updateSyncStatus('從 Sheet 載入中…');
  const ok = await syncFromSheet();
  if (ok) {
    // ── 完整重建所有 UI ──
    // 1. 旅程標題 & 日期（桌面 + 手機 header）
    $('sidebar-theme').textContent = S.trip.theme;
    $('sidebar-dates').textContent = `${fmtShort(S.trip.start)} — ${fmtShort(S.trip.end)}`;
    if ($('mobile-theme')) $('mobile-theme').textContent = S.trip.theme;
    if ($('mobile-dates')) $('mobile-dates').textContent = `${fmtShort(S.trip.start)} — ${fmtShort(S.trip.end)}`;
    if ($('search-theme-hint')) $('search-theme-hint').textContent = S.trip.theme;
    // 2. 日期列表（桌面 sidebar + 手機橫向捲軸）
    buildDateList();
    // 3. 選取日期（保留當前選取，如超出範圍則切回第一天）
    const d = (S.selectedDate && S.selectedDate >= S.trip.start && S.selectedDate <= S.trip.end)
      ? S.selectedDate : getDefaultDate();
    selectDate(d);  // 同時重繪行程列表
    // 4. 其他面板
    renderFlights();
    renderAccom();
    renderExpenses();
    renderMemos();
    // 5. 地圖重置（避免顯示已被刪除的舊地點）
    // selectedItemId already cleared in syncFromSheet if item was deleted
    resetMap();
    // 6. sync bar
    updateSyncUI();
    toast('✓ 已從 Google Sheet 拉取最新資料');
    scheduleAutoRefresh();
  } else {
    updateSyncStatus('⚠ 無法從 Sheet 讀取（確認 GAS 與 Sheet 設定）');
  }
}

// ── 協作者心跳 & 人數 ────────────────
// 每 30 秒向 GAS 發送 heartbeat，讀回已上線組員數
const S_SESSION = uid(); // 本安裝對应的唯一 session ID
let S_COLLAB_TIMER = null;

async function sendHeartbeat() {
  if (!S.gasUrl || !S.sheetId) return;
  try {
    const payload = {
      id: S.sheetId,
      action: 'heartbeat',
      session: S_SESSION,
      name: S.trip.theme || '團隊成員',
    };
    // no-cors mode 無法讀回 response，改用 GET 方式担心跟蹤
    const resp = await fetch(`${S.gasUrl}?id=${S.sheetId}&action=heartbeat&session=${S_SESSION}&ts=${Date.now()}`);
    const text = await resp.text();
    const d = JSON.parse(text);
    if (d.count !== undefined) updateCollabCount(d.count);
  } catch (e) { /* 静默失敗，不影鞿主功能 */ }
}

function updateCollabCount(count) {
  let badge = document.getElementById('collab-badge');
  if (!badge) return;
  if (count > 0) {
    badge.textContent = `👥 ${count} 人協作中`;
    badge.style.display = 'inline-flex';
  } else {
    badge.style.display = 'none';
  }
}

function updateCollabLocal() {
  // 沒有 GAS 時，顯示本機單人狀態
  let badge = document.getElementById('collab-badge');
  if (!badge) return;
  if (S.sheetId && !S.gasUrl) {
    badge.textContent = '👤 1 人（本機）';
    badge.style.display = 'inline-flex';
  } else if (!S.sheetId) {
    badge.style.display = 'none';
  }
}

function startHeartbeat() {
  if (S_COLLAB_TIMER) clearInterval(S_COLLAB_TIMER);
  if (!S.gasUrl || !S.sheetId) {
    // 沒有 GAS，仍顯示本機單人狀態
    updateCollabLocal();
    return;
  }
  sendHeartbeat(); // 立即發送一次
  S_COLLAB_TIMER = setInterval(sendHeartbeat, 30000); // 每 30 秒
}

// ── DEMO DATA ─────────────────────────────
function loadDemo() {
  S.itinerary = [
    { id: uid(), name: '博多拉麵一蘭本店', addr: '福岡市博多區中洲5-3-2', query: '博多拉麵一蘭 福岡', gmaplink: '', date: '2025-03-15', time: '18:00', duration: 90, type: 'restaurant', note: '招牌豚骨拉麵，記得選靠窗位置', expenses: [], travelTime: null },
    { id: uid(), name: 'Canal City 博多', addr: '福岡市博多區住吉1-2', query: 'Canal City 博多', gmaplink: '', date: '2025-03-15', time: '20:00', duration: 60, type: 'shopping', note: '大型購物中心，夜間噴水秀', expenses: [], travelTime: null },
    { id: uid(), name: '太宰府天滿宮', addr: '太宰府市宰府4-7-1', query: '太宰府天滿宮 福岡', gmaplink: '', date: '2025-03-16', time: '09:00', duration: 120, type: 'attraction', note: '必買梅枝餅！', expenses: [], travelTime: null },
    { id: uid(), name: '隈研吾設計星巴克', addr: '太宰府市宰府3-2-43', query: '隈研吾 星巴克 太宰府', gmaplink: '', date: '2025-03-16', time: '11:30', duration: 45, type: 'restaurant', note: '木條格柵建築，必打卡', expenses: [], travelTime: null },
    { id: uid(), name: '長崎平和公園', addr: '長崎市松山町', query: '長崎平和公園', gmaplink: '', date: '2025-03-17', time: '10:00', duration: 90, type: 'attraction', note: '和平紀念像', expenses: [], travelTime: null },
    { id: uid(), name: '哥拉巴園', addr: '長崎市南山手町8-1', query: '哥拉巴園 長崎', gmaplink: '', date: '2025-03-17', time: '14:00', duration: 90, type: 'attraction', note: '俯瞰長崎港灣美景', expenses: [], travelTime: null },
    { id: uid(), name: '別府地獄溫泉', addr: '別府市鐵輪', query: '別府地獄溫泉 大分', gmaplink: '', date: '2025-03-18', time: '09:30', duration: 180, type: 'attraction', note: '血の池地獄必去', expenses: [], travelTime: null },
    { id: uid(), name: '由布院金鱗湖', addr: '由布市湯布院町川上', query: '由布院金鱗湖 大分', gmaplink: '', date: '2025-03-19', time: '08:00', duration: 120, type: 'attraction', note: '早晨薄霧最美', expenses: [], travelTime: null },
  ];
  S.flights = [
    { id: uid(), number: 'JL 0931', from: 'TPE', to: 'FUK', depart: '2025-03-15T10:00', arrive: '2025-03-15T13:40', note: '桃園T2' },
    { id: uid(), number: 'JL 0932', from: 'FUK', to: 'TPE', depart: '2025-03-22T14:30', arrive: '2025-03-22T16:20', note: '福岡T2' },
  ];
  S.accommodation = [
    { id: uid(), name: '博多Canal City Washington Hotel', addr: '福岡市博多區住吉1-2-20', checkin: '2025-03-15', checkout: '2025-03-17', note: '雙人標準房', gmaplink: '' },
    { id: uid(), name: '由布院溫泉旅館', addr: '大分縣由布市湯布院町', checkin: '2025-03-19', checkout: '2025-03-20', note: '附早晚餐', gmaplink: '' },
  ];
  S.expenses = [
    { id: uid(), date: '2025-03-15', category: 'food', desc: '博多拉麵一蘭', amount: 980, currency: 'JPY', payer: '小明', splits: [], linkedItemId: '' },
    { id: uid(), date: '2025-03-15', category: 'transport', desc: '機場巴士', amount: 620, currency: 'JPY', payer: '小美', splits: [], linkedItemId: '' },
    { id: uid(), date: '2025-03-16', category: 'food', desc: '梅枝餅', amount: 800, currency: 'JPY', payer: '小明', splits: [], linkedItemId: '' },
  ];
  S.history = [{ t: nowStr(), a: '載入範例資料（日本九州）' }];
}

// ── MEMOS ─────────────────────────────────
function renderMemos() {
  const list = $('memo-list');
  if (!list) return;
  if (!S.memos.length) {
    list.innerHTML = '';
    const empty = document.createElement('div');
    empty.className = 'empty-state';
    empty.id = 'memo-empty';
    empty.innerHTML = '<div class="empty-icon">✎</div><p>尚無備忘錄<br><span>點擊右上角新增主題與內容</span></p>';
    list.appendChild(empty);
    return;
  }
  list.innerHTML = S.memos.map(m => `
    <div class="memo-card" id="memo-card-${m.id}">
      <div class="memo-card-header">
        <div class="memo-card-title">${escapeHtml(m.title)}</div>
        <div class="memo-card-actions">
          <button class="item-action-btn" onclick="openEditMemoModal('${m.id}')">編輯</button>
          <button class="item-action-btn del-btn" onclick="deleteMemo('${m.id}')">刪除</button>
        </div>
      </div>
      <div class="memo-card-content">${escapeHtml(m.content).replace(/\n/g, '<br>')}</div>
      <div class="memo-card-date">${m.updatedAt || ''}</div>
    </div>
  `).join('');
}

function escapeHtml(str) {
  return (str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function openAddMemoModal() {
  $('memo-modal-title').textContent = '新增備忘錄';
  $('memo-edit-id').value = '';
  $('memo-title-input').value = '';
  $('memo-content-input').value = '';
  openModal('modal-memo');
}

function openEditMemoModal(id) {
  const m = S.memos.find(x => x.id === id);
  if (!m) return;
  $('memo-modal-title').textContent = '編輯備忘錄';
  $('memo-edit-id').value = id;
  $('memo-title-input').value = m.title;
  $('memo-content-input').value = m.content;
  openModal('modal-memo');
}

function saveMemoModal() {
  const title = $('memo-title-input').value.trim();
  const content = $('memo-content-input').value.trim();
  if (!title) { toast('請填寫備忘錄主題'); return; }
  const editId = $('memo-edit-id').value;
  if (editId) {
    const m = S.memos.find(x => x.id === editId);
    if (m) { m.title = title; m.content = content; m.updatedAt = nowStr(); }
    toast('備忘錄已更新');
    log('更新備忘錄：' + title);
  } else {
    S.memos.push({ id: uid(), title, content, updatedAt: nowStr() });
    toast('備忘錄已儲存：' + title);
    log('新增備忘錄：' + title);
  }
  renderMemos();
  closeModal('modal-memo');
  scheduleSave();
}

function deleteMemo(id) {
  const m = S.memos.find(x => x.id === id);
  if (!m) return;
  showDeleteConfirm(m.title, () => {
    S.memos = S.memos.filter(x => x.id !== id);
    renderMemos();
    toast('已刪除備忘錄：' + m.title);
    log('刪除備忘錄：' + m.title);
    scheduleSave();
  });
}

// ── MODAL HELPERS ─────────────────────────
function openModal(id) { $(id).classList.add('active'); }
function closeModal(id) { $(id).classList.remove('active'); }

document.addEventListener('click', e => {
  if (e.target.classList.contains('modal-overlay') && e.target.id !== 'setup-modal')
    e.target.classList.remove('active');
});

// ── MOBILE BOTTOM NAV ─────────────────────
function switchPanelMobile(panelName) {
  document.querySelectorAll('.nav-item').forEach(b => b.classList.toggle('active', b.dataset.panel === panelName));
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  const target = $('panel-' + panelName);
  if (target) target.classList.add('active');
  document.querySelectorAll('.mobile-nav-btn').forEach(b => b.classList.toggle('active', b.dataset.panel === panelName));
}

// （utility 函式已移至檔案頂端）

// ── RESPONSIVE MAP REBUILD ────────────────
// 手機轉向或視窗大小改變時，確保 mobile map placeholder 已建立
window.addEventListener('orientationchange', () => {
  setTimeout(() => {
    setupMobileMap();
    // 如果地圖正在顯示，重整 iframe src 以觸發重繪
    const iframe = $('mobile-map-iframe');
    if (iframe && iframe.src && iframe.src !== window.location.href) {
      const src = iframe.src;
      iframe.src = '';
      setTimeout(() => { iframe.src = src; }, 50);
    }
  }, 300);
});