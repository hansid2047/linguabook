// ═══════════════════════════════════════════════════════════════
// LinguaBook Pro · Google Apps Script 後端程式碼
// ───────────────────────────────────────────────────────────────
// 使用方法：
//   1. 開啟你的 Google 試算表
//   2. 點選選單「擴充功能」→「Apps Script」
//   3. 把編輯器裡原有的程式碼全部刪掉
//   4. 把這個檔案的內容全部複製，貼進去
//   5. 按 Ctrl + S 儲存（專案名稱隨意填，例如 LinguaBook）
//   6. 點右上角「部署」→「新增部署作業」
//      ▸ 類型：網頁應用程式
//      ▸ 執行身分：我
//      ▸ 存取權限：任何人   ← 這個很重要！
//   7. 點「部署」→ 複製產生的「網頁應用程式網址」
//   8. 把網址貼到 LinguaBook 的「設定」頁面即可
// ═══════════════════════════════════════════════════════════════

// ── 設定：工作表名稱（不需要改） ──────────────────────────────
const SHEET_NAME = 'Books';

// ── 取得工作表，不存在就自動建立 ──────────────────────────────
function getSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    // 自動建立工作表並加入標題列
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'id', 'title', 'author', 'level', 'type',
      'emoji', 'color', 'text', 'desc', 'createdAt'
    ]);
    // 美化標題列
    sheet.getRange(1, 1, 1, 10)
      .setFontWeight('bold')
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(8, 500); // text 欄加寬方便檢視
  }

  return sheet;
}

// ── 處理 GET 請求（讀取資料） ──────────────────────────────────
// 當 LinguaBook 網頁讀取書庫時，會呼叫這裡
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || 'list';

    if (action === 'ping') {
      // 測試連線用
      return ok({ message: 'LinguaBook 連線成功 ✅' });
    }

    if (action === 'list') {
      // 回傳所有書本
      return ok({ books: getAllBooks() });
    }

    return err('未知的 action：' + action);

  } catch (e) {
    return err(e.toString());
  }
}

// ── 處理 POST 請求（寫入 / 刪除資料） ─────────────────────────
// 當 LinguaBook 新增或刪除書本時，會呼叫這裡
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'add') {
      return ok(addBook(body.book));
    }

    if (action === 'delete') {
      return ok(deleteBook(body.id));
    }

    return err('未知的 action：' + action);

  } catch (e) {
    return err(e.toString());
  }
}

// ── 讀取所有書本 ───────────────────────────────────────────────
function getAllBooks() {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();

  // 第一列是標題，從第二列開始讀
  return rows.slice(1)
    .filter(r => r[0]) // 過濾空列
    .map(r => ({
      id:        r[0],
      title:     r[1],
      author:    r[2],
      level:     r[3],
      type:      r[4],
      emoji:     r[5] || '📖',
      color:     r[6] || '#f0c040',
      text:      r[7],
      desc:      r[8] || '',
      createdAt: r[9]
    }));
}

// ── 新增書本 ───────────────────────────────────────────────────
function addBook(b) {
  const sheet = getSheet();
  const id    = 'gs_' + Date.now(); // 用時間戳產生唯一 ID

  sheet.appendRow([
    id,
    b.title  || '未命名',
    b.author || '',
    b.level  || 'B1',
    b.type   || 'article',
    b.emoji  || '📖',
    b.color  || '#f0c040',
    b.text   || '',
    b.desc   || '',
    new Date().toISOString()
  ]);

  return {
    id:      id,
    message: '書本已儲存到 Google Sheets ✅'
  };
}

// ── 刪除書本 ───────────────────────────────────────────────────
function deleteBook(id) {
  const sheet = getSheet();
  const data  = sheet.getDataRange().getValues();

  // 從第二列開始找符合 id 的那一列
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1); // Sheet 列號從 1 開始，資料從第 2 列
      return { message: '書本已刪除 ✅' };
    }
  }

  throw new Error('找不到書本，id = ' + id);
}

// ── 回應格式工具函式 ───────────────────────────────────────────
function ok(data) {
  return out(Object.assign({ ok: true }, data));
}

function err(message) {
  return out({ ok: false, error: message });
}

function out(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
