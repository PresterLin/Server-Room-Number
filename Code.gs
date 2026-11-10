/**
 * ============================================================
 *  樓層導覽系統（博物館式）— 伺服器端程式 Code.gs
 * ============================================================
 *  功能：
 *   1. doGet()         ：提供網頁入口（部署為 Web App 後開啟的網址）
 *   2. getInitData()   ：讀取 Google 試算表「房號房名」A:D，回傳給前端畫樹狀選單
 *   3. getFloorSvg()   ：依樓層名稱，回傳對應的 SVG 圖層（存放在專案的 SVG_xxx.html 檔）
 *
 *  專案檔案結構（Apps Script 編輯器左側「檔案」）：
 *   ├─ Code.gs        （本檔，類型：指令碼）
 *   ├─ Index.html     （前端網頁，類型：HTML）
 *   ├─ SVG_B1F.html   （B1F 圖層，類型：HTML）
 *   ├─ SVG_1F.html    （1F 圖層）
 *   ├─ SVG_2F.html … SVG_5F.html、SVG_R.html（屋頂層）
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

/** 網頁標題 */
const APP_TITLE = '台電彰化雲端資料中心｜樓層導覽';

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
  return { floors: floors, rooms: rooms };
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
 * 供前端把 HTML 片段檔嵌入主頁的小工具（本專案目前未使用，保留給日後拆分 CSS/JS 用）。
 * 用法：在 Index.html 寫 <?!= include('Style') ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
