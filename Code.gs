/**
 * 1277 Food — Apps Script (Code.gs)
 *
 * CÁCH DÙNG
 * 1) Tạo Google Sheets mới -> Tiện ích mở rộng -> Apps Script (khuyên dùng, để Sheet gắn liền với mã).
 *    Hoặc tạo dự án Apps Script độc lập: mã sẽ tự tạo 1 Google Sheets "1277 Food - Luot xem".
 * 2) Dán toàn bộ file này vào Code.gs.
 * 3) Thêm file HTML tên đúng là "index" (Tệp -> HTML) và dán toàn bộ nội dung index.html.
 * 4) Chọn hàm "setup" -> Chạy 1 lần để cấp quyền (tạo sẵn trang tính lượt xem).
 * 5) Triển khai -> Tùy chọn triển khai mới -> Ứng dụng web:
 *      Thực thi với tư cách: Tôi   |   Ai có quyền truy cập: Bất kỳ ai
 *    Mở link /exec là dùng được.
 *
 * Lượt xem: mỗi trình duyệt chỉ tính 1 lần (theo mã thiết bị lưu trong trình duyệt).
 * Dữ liệu lượt xem nằm ở trang tính "LuotXem" (cột A = mã thiết bị, B = lần đầu xem, C = trình duyệt).
 */

var SHEET_NAME = 'LuotXem';
var PROP_SHEET_ID = 'VIEWS_SPREADSHEET_ID';
var CACHE_KEY = 'views_total_v1';

/** Trang chính + API JSON (cho bản host ngoài như GitHub Pages). */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action) {
    return handleApi_(p);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('1277 Food')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Chạy 1 lần trong trình soạn thảo để cấp quyền và tạo sẵn trang tính. */
function setup() {
  var sh = getSheet_();
  Logger.log('Sẵn sàng. Trang tính: ' + sh.getParent().getUrl());
  return sh.getParent().getUrl();
}

/* ====== Hàm gọi từ giao diện (google.script.run) ====== */

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

function getViews() {
  return { ok: true, total: readTotal_() };
}

function registerView(deviceId, userAgent) {
  return registerView_(deviceId, userAgent);
}

/* ====== API JSON (GET ?action=views | ?action=view&id=...) ====== */

function handleApi_(p) {
  var out;
  try {
    if (p.action === 'views') {
      out = { ok: true, total: readTotal_() };
    } else if (p.action === 'view') {
      out = registerView_(p.id, p.ua || '');
    } else {
      out = { ok: false, error: 'unknown_action' };
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ====== Xử lý lượt xem ====== */

function registerView_(deviceId, userAgent) {
  var id = String(deviceId || '');
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(id)) {
    return { ok: false, error: 'bad_id', total: readTotal_() };
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = getSheet_();
    var last = sh.getLastRow();
    var isNew = true;
    if (last >= 2) {
      var found = sh.getRange(2, 1, last - 1, 1)
        .createTextFinder(id)
        .matchEntireCell(true)
        .matchCase(true)
        .findNext();
      isNew = !found;
    }
    if (isNew) {
      sh.appendRow([id, new Date(), String(userAgent || '').slice(0, 200)]);
      SpreadsheetApp.flush();
    }
    var total = Math.max(0, sh.getLastRow() - 1);
    CacheService.getScriptCache().put(CACHE_KEY, String(total), 20);
    return { ok: true, total: total, isNew: isNew };
  } finally {
    lock.releaseLock();
  }
}

function readTotal_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(CACHE_KEY);
  if (hit !== null && hit !== undefined && !isNaN(Number(hit))) {
    return Number(hit);
  }
  var total = Math.max(0, getSheet_().getLastRow() - 1);
  cache.put(CACHE_KEY, String(total), 20);
  return total;
}

function getSheet_() {
  var ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    ss = null;
  }
  if (!ss) {
    var props = PropertiesService.getScriptProperties();
    var savedId = props.getProperty(PROP_SHEET_ID);
    if (savedId) {
      try {
        ss = SpreadsheetApp.openById(savedId);
      } catch (err2) {
        ss = null;
      }
    }
    if (!ss) {
      ss = SpreadsheetApp.create('1277 Food - Luot xem');
      props.setProperty(PROP_SHEET_ID, ss.getId());
    }
  }
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 3).setValues([['Ma thiet bi', 'Lan dau xem', 'Trinh duyet']]);
    sh.setFrozenRows(1);
  }
  return sh;
}
