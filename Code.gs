/**
 * ============================================================
 *  樓層導覽系統（博物館式）— 伺服器端程式 Code.gs
 * ============================================================
 *  功能：
 *   1. doGet()         ：提供網頁入口（部署為 Web App 後開啟的網址）
 *   2. getInitData()   ：讀取 Google 試算表「房號房名」A:D，回傳給前端畫樹狀選單
 *   3. getFloorSvg()   ：依樓層名稱，回傳對應的 SVG 圖層（存放在專案的 SVG_xxx.html 檔）
 *   4. getFloorImage() ：依樓層名稱，從 Google Drive 資料夾讀取 3D 轉 2D 的彩色 PNG 底圖，
 *                        轉成 data URL 回傳給網頁（不需把圖檔設成公開分享）
 *
 *  專案檔案結構（Apps Script 編輯器左側「檔案」）：
 *   ├─ Code.gs        （本檔，類型：指令碼）
 *   ├─ Index.html     （前端網頁，類型：HTML）
 *   ├─ SVG_B1F.html   （B1F 圖層，類型：HTML）
 *   ├─ SVG_1F.html    （1F 圖層）
 *   ├─ SVG_2F.html … SVG_5F.html、SVG_R.html（屋頂層）
 *   （PNG 底圖不在專案內，放在 Drive 的 IMG_FOLDER_ID 資料夾，檔名見 FLOOR_IMG_MAP）
 *
 *  注意：Apps Script 的「HTML」檔在編輯器中新增時，名稱不必加 .html（系統會自動加）。
 */

// ───────────────────────── 設定區（只需要改這裡） ─────────────────────────

/** 試算表 ID（網址 /d/ 與 /edit 之間那一串） */
const SHEET_ID = '1-gVgdHREKTWIBzkY06uLN7Pq4KpocfA8hUWrfC_S-us';

/** 工作表名稱 */
const SHEET_NAME = '房號房名';

/**
 * 【樓層對照表】：試算表「樓層」欄的文字  →  專案內 SVG 檔名（不含 .html）
 * 例如：試算表 A 欄寫「1F」，系統就會載入 SVG_1F.html。
 * 新增樓層時：① 新增一個 SVG_xxx 檔  ② 在這裡加一行  ③ 試算表「樓層」欄填入同樣的文字。
 */
const FLOOR_MAP = {
  'B1F': 'SVG_B1F',
  '1F': 'SVG_1F',
  '2F': 'SVG_2F',
  '3F': 'SVG_3F',
  '4F': 'SVG_4F',
  '5F': 'SVG_5F',
  '屋頂層': 'SVG_R'
};

/**
 * 【底圖資料夾】存放 3D 轉 2D 彩色 PNG 底圖的 Google Drive 資料夾 ID
 * （資料夾網址 /folders/ 後面那一串）。Web App 以「我」的身分執行時，使用者不需要有此資料夾的權限。
 */
const IMG_FOLDER_ID = '1DIDm7i0jXFPo5N9Rv-Gz2p3vQxUNdZka';

/**
 * 【底圖檔名對照表】試算表「樓層」欄的文字  →  Drive 資料夾內的 PNG 檔名（需完全一致，含副檔名）
 * 沒有列在這裡的樓層（或 Drive 找不到檔案）時，網頁會自動退回「牆線圖」，不會當機。
 */
const FLOOR_IMG_MAP = {
  'B1F': 'B1F.png',
  '1F': '1F.png',
  '2F': '2F.png',
  '3F': '3F.png',
  '4F': '4F.png',
  '5F': '5F.png',
  '屋頂層': '屋頂層.png'
};

/**
 * 【底圖連結來源】工作表「統計」的 D:F 欄（D=樓層、F=彩色底圖的 Google Drive 連結）。
 * 網頁會直接用這些連結的「檔案 ID」從 Google 伺服器載入底圖，比經過 Apps Script 轉傳更快。
 * 讀不到或載入失敗時，會自動退回原本的 getFloorImage()（資料夾 + base64）方式。
 */
const LINK_SHEET_NAME = '統計';

/** 網頁標題 */
const APP_TITLE = '華陽備勤大樓｜樓層導覽';

// ───────────────────────── 網頁入口 ─────────────────────────

/**
 * 使用者開啟 Web App 網址時，Apps Script 會自動呼叫此函式。
 * 回傳 Index.html 的內容（使用樣板，方便未來傳入參數）。
 */
function doGet() {
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.appTitle = APP_TITLE;                       // 傳給 Index.html 的 <?= appTitle ?>
  return tpl.evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ───────────────────────── 提供資料給前端 ─────────────────────────

/**
 * 讀取試算表資料，整理成前端需要的格式。
 * 回傳：{ floors: ['B1F','1F',...], rooms: [{floor, room, name, note, hasSvg}, ...] }
 *  - floors 的順序＝試算表中樓層「第一次出現」的順序
 *  - 房號一律轉成文字並去除前後空白，避免 101 變成數字 101、或多餘空格造成比對失敗
 */
function getInitData() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('找不到工作表：' + SHEET_NAME);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { floors: [], rooms: [] };

  // 讀取 A:D（第 2 列起，跳過標題列）。getDisplayValues 取得「畫面上看到的文字」，
  // 可避免房號被 Google 試算表自動轉成數字。
  const values = sheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues();

  const floors = [];
  const rooms = [];
  values.forEach(function (r) {
    const floor = String(r[0]).trim();
    const room = String(r[1]).trim();
    if (!floor || !room) return;                  // 空白列略過
    if (floors.indexOf(floor) === -1) floors.push(floor);
    rooms.push({
      floor: floor,
      room: room,
      name: String(r[2]).trim(),                  // 房名可為空白（例如空調管道）
      note: String(r[3]).trim(),
      hasSvg: !!FLOOR_MAP[floor]                  // 該樓層是否有對應圖層
    });
  });
  return { floors: floors, rooms: rooms, imgIds: getImageIds_() };
}

/**
 * 從「統計」工作表 D:F 讀出 { 樓層: Drive 檔案 ID }。
 * F 欄可以是純文字網址，也可以是「文字加超連結」；支援 /file/d/ID/view 與 ?id=ID 兩種網址格式，也接受直接貼 ID。
 * 工作表不存在或格式不對時回傳空物件（網頁會改走舊方式，不會當機）。
 */
function getImageIds_() {
  const ids = {};
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LINK_SHEET_NAME);
    if (!sh || sh.getLastRow() < 2) return ids;
    const rng = sh.getRange(2, 4, sh.getLastRow() - 1, 3);          // 第 2 列起的 D:F
    const vals = rng.getDisplayValues();
    const rich = rng.getRichTextValues();
    for (let i = 0; i < vals.length; i++) {
      const floor = String(vals[i][0]).trim();
      if (!floor) continue;
      const src = rich[i][2].getLinkUrl() || String(vals[i][2]).trim();   // 優先取超連結，否則用儲存格文字
      const m = src.match(/\/d\/([A-Za-z0-9_-]{15,})/) || src.match(/[?&]id=([A-Za-z0-9_-]{15,})/) || src.match(/^([A-Za-z0-9_-]{25,})$/);
      if (m) ids[floor] = m[1];
    }
  } catch (e) { /* 讀取失敗就視為沒有連結 */ }
  return ids;
}

/**
 * 依樓層名稱回傳 SVG 字串。前端切換樓層時才呼叫（延遲載入，加快首頁速度）。
 * 找不到對應檔案時回傳 null，前端會顯示「尚無圖層」。
 */
function getFloorSvg(floor) {
  const file = FLOOR_MAP[floor];
  if (!file) return null;
  try {
    // createHtmlOutputFromFile 讀取專案內 HTML 檔，getContent() 取得純文字內容
    return HtmlService.createHtmlOutputFromFile(file).getContent();
  } catch (e) {
    return null;                                  // 檔案不存在
  }
}

/**
 * 在底圖資料夾中尋找指定檔名的圖檔（比對方式比 getFilesByName 寬鬆、也能回報原因）。
 *  - 全形／半形、大小寫、前後或中間多餘空白都視為相同（例如「1F.png」與「1F.png 」「１Ｆ.PNG」）
 *  - 同名檔案有多個時，取「最後修改」的那一個
 *  - 若找到的是 Drive「捷徑」，自動改讀捷徑指向的原始檔案
 * 回傳：{ file: 找到的檔案或 null, count: 符合的檔案數, names: 資料夾內所有檔名（找不到時用來提示） }
 */
function findImageFile_(name) {
  const norm = function (t) { return String(t).normalize('NFKC').replace(/\s+/g, '').toLowerCase(); };
  const want = norm(name);
  const wantBase = want.replace(/\.[^.]+$/, '');                 // 不含副檔名的部分
  const it = DriveApp.getFolderById(IMG_FOLDER_ID).getFiles();
  const hits = [], names = [];
  while (it.hasNext()) {
    const f = it.next();
    const n = norm(f.getName());
    names.push(f.getName());
    if (n === want || n.replace(/\.[^.]+$/, '') === wantBase) hits.push(f);
  }
  hits.sort(function (x, y) { return y.getLastUpdated() - x.getLastUpdated(); });   // 最新的排前面
  let file = hits.length ? hits[0] : null;
  if (file && file.getMimeType() === 'application/vnd.google-apps.shortcut') {
    file = DriveApp.getFileById(file.getTargetId());              // 捷徑 → 原始檔
  }
  return { file: file, count: hits.length, names: names };
}

/**
 * 依樓層名稱，從 Drive 資料夾讀取 PNG 底圖。
 * 回傳物件：{ url: 'data:image/png;base64,…' } 成功；{ url: null, msg: '失敗原因' } 失敗。
 * 失敗原因會顯示在網頁上，方便找出問題（找不到檔、沒有權限、檔案損毀…）。
 * 注意：第一次執行會要求「Google Drive」授權，請按「允許」。
 */
function getFloorImage(floor) {
  const name = FLOOR_IMG_MAP[floor];
  if (!name) return { url: null, msg: 'FLOOR_IMG_MAP 沒有設定「' + floor + '」的檔名' };
  try {
    const r = findImageFile_(name);
    if (!r.file) {
      return { url: null, msg: '資料夾內找不到「' + name + '」。目前資料夾的檔案：' + r.names.join('、') };
    }
    const blob = r.file.getBlob();
    let type = blob.getContentType();
    if (!type || type.indexOf('image/') !== 0) {                  // Drive 記錄的類型不是圖片時，依副檔名推測
      type = /\.jpe?g$/i.test(r.file.getName()) ? 'image/jpeg' : 'image/png';
    }
    return { url: 'data:' + type + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
  } catch (e) {
    return { url: null, msg: '讀取「' + name + '」失敗：' + e.message };
  }
}

/**
 * 【診斷工具】在 Apps Script 編輯器上方選擇函式「diagnoseImages」按「執行」，
 * 再看下方「執行記錄」，會逐層列出：找到幾個符合的檔案、檔案 ID、類型、大小、最後修改時間、是否讀取成功。
 */
function diagnoseImages() {
  Object.keys(FLOOR_IMG_MAP).forEach(function (fl) {
    const name = FLOOR_IMG_MAP[fl];
    try {
      const r = findImageFile_(name);
      if (!r.file) { Logger.log('【' + fl + '】找不到「' + name + '」。資料夾內檔案：' + r.names.join(' | ')); return; }
      const f = r.file, bytes = f.getBlob().getBytes().length;
      Logger.log('【' + fl + '】OK　符合檔案數=' + r.count + '　名稱=「' + f.getName() + '」(長度 ' + f.getName().length +
        ')　ID=' + f.getId() + '　類型=' + f.getMimeType() + '　大小=' + Math.round(bytes / 1024) + 'KB　修改=' + f.getLastUpdated());
    } catch (e) {
      Logger.log('【' + fl + '】失敗：' + e.message);
    }
  });
}

/**
 * 供前端把 HTML 片段檔嵌入主頁的小工具（本專案目前未使用，保留給日後拆分 CSS/JS 用）。
 * 用法：在 Index.html 寫 <?!= include('Style') ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
