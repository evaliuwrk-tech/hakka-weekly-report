/**
 * 哈客組 週報填報系統 — Google Apps Script 後端
 * ------------------------------------------------------------
 * 功能：
 *   1. doPost：成員按下「產生週報」時，把繳交紀錄寫進試算表（同一人同一週會覆蓋，不重複）
 *   2. doGet ：進度看板用 JSONP 讀取「誰在哪一週交了」（只回傳狀態，不回傳內文）
 *
 * 部署步驟見「設定說明.md」。部署後把網址貼進 index.html 的 CONFIG.appsScriptUrl。
 */

var SHEET_NAME = 'Submissions';
var HEADERS = ['提交時間', '週次', '週次起訖', '姓名', '本週進度', '困難與處理', '下週規劃', '備註'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 成員送出：新增或覆蓋（依「週次＋姓名」判斷） */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    var data = JSON.parse(e.postData.contents);
    var sh = getSheet_();
    var row = [
      new Date(),
      data.week || '',
      data.weekRange || '',
      data.name || '',
      data.progress || '',
      data.difficulties || '',
      data.nextPlan || '',
      data.note || ''
    ];

    // 找同一週、同一人的既有列 → 覆蓋
    var last = sh.getLastRow();
    var foundRow = 0;
    if (last > 1) {
      var vals = sh.getRange(2, 2, last - 1, 3).getValues(); // 週次, 週次起訖, 姓名
      for (var i = 0; i < vals.length; i++) {
        if (String(vals[i][0]) === String(data.week) && String(vals[i][2]) === String(data.name)) {
          foundRow = i + 2;
          break;
        }
      }
    }
    if (foundRow) {
      sh.getRange(foundRow, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** 看板讀取：只回傳「時間 / 週次 / 姓名」，不含週報內文 */
function doGet(e) {
  var rows = [];
  try {
    var sh = getSheet_();
    var last = sh.getLastRow();
    if (last > 1) {
      var vals = sh.getRange(2, 1, last - 1, 4).getValues(); // 提交時間,週次,起訖,姓名
      rows = vals.map(function (r) {
        return {
          time: r[0] ? Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'MM/dd HH:mm') : '',
          week: r[1],
          weekRange: r[2],
          name: r[3]
        };
      });
    }
  } catch (err) {
    rows = [];
  }
  var out = JSON.stringify({ ok: true, rows: rows });
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + out + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
