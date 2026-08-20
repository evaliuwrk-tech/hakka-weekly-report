/**
 * 週報填報系統 — Google Apps Script 後端（多組版）
 * ------------------------------------------------------------
 *  1. doPost：(a) 網頁填報 → 依「週次+姓名」寫入/覆蓋（含組別欄）
 *            (b) LINE Webhook → 把最近收到的群組 groupId 存進 LAST_GROUP_ID
 *  2. doGet ：進度看板用 JSONP 讀取繳交狀態（不含內文）
 *  3. sendOverdueReminder：每週五 16:30 觸發器自動跑，
 *       逐組把「本週逾期未交」名單推播到各組自己的 LINE 群組
 *
 *  ★ 改動此檔後，務必重新部署（管理部署作業→編輯✏️→版本選「新版本」→部署），網址不變。
 *  ★ GROUPS 需與 config.js 的 groups 保持一致。
 */

/* ===== 需與 config.js 一致 ===== */
var FORM_URL = 'https://evaliuwrk-tech.github.io/hakka-weekly-report/';
var GROUPS = [
  { key: 'hakka', name: '哈客組',
    members: ['劉怡君（組長）', '涂政強', '黃敏琪', '劉美君'] },
  { key: 'lsk',   name: '來上客組',
    members: ['邱嘉圓（組長）', '吳欣蓉（葉子）', '林郁婷', '蔡旻宏（阿蔡）'] }
];

var SHEET_NAME = 'Submissions';
var HEADERS = ['提交時間', '週次', '週次起訖', '姓名', '本週進度', '困難與處理', '下週規劃', '備註', '組別'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  return sh;
}

/* ============================ doPost ============================ */
function doPost(e) {
  var data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return ok_('bad'); }

  // (b) LINE Webhook 事件 → 記錄最近一次群組 ID
  if (data && data.events) {
    try {
      var props = PropertiesService.getScriptProperties();
      data.events.forEach(function (ev) {
        var src = (ev && ev.source) || {};
        var id = src.groupId || src.roomId;
        if (id) props.setProperty('LAST_GROUP_ID', id);
      });
    } catch (err2) {}
    return ContentService.createTextOutput('OK');
  }

  // (a) 網頁填報 → upsert（依 週次+姓名）
  var lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    var sh = getSheet_();
    var row = [ new Date(), data.week || '', data.weekRange || '', data.name || '',
                data.progress || '', data.difficulties || '', data.nextPlan || '', data.note || '',
                data.groupName || '' ];
    var last = sh.getLastRow(), foundRow = 0;
    if (last > 1) {
      var vals = sh.getRange(2, 2, last - 1, 3).getValues(); // 週次, 起訖, 姓名
      for (var i = 0; i < vals.length; i++) {
        if (String(vals[i][0]) === String(data.week) && String(vals[i][2]) === String(data.name)) { foundRow = i + 2; break; }
      }
    }
    if (foundRow) sh.getRange(foundRow, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally { lock.releaseLock(); }
}

/* ============================ doGet（看板） ============================ */
function doGet(e) {
  var rows = [];
  try {
    var sh = getSheet_(), last = sh.getLastRow();
    if (last > 1) {
      var vals = sh.getRange(2, 1, last - 1, 9).getValues(); // 時間,週次,起訖,姓名,...,組別
      rows = vals.map(function (r) {
        return {
          time: r[0] ? Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'MM/dd HH:mm') : '',
          week: r[1], weekRange: r[2], name: r[3], group: r[8] || ''
        };
      });
    }
  } catch (err) { rows = []; }
  var out = JSON.stringify({ ok: true, ver: 'multi-1', rows: rows });
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) return ContentService.createTextOutput(cb + '(' + out + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

/* ============================ LINE 逾期催繳（逐組） ============================ */
function sendOverdueReminder() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('LINE_TOKEN');
  if (!token) return;
  GROUPS.forEach(function (g) {
    var gid = props.getProperty('LINE_GROUP_ID_' + g.key);
    if (!gid) return;                       // 該組沒綁定群組就略過
    var info = missingForGroup_(g);
    if (!info.missing.length) return;       // 全繳齊就不打擾
    var msg = '【' + g.name + ' 週報逾期提醒】\n'
      + '本週（' + info.range + '）週報已過截止（週五 16:00），以下同仁尚未繳交，請儘快補上，謝謝！\n'
      + info.missing.map(function (m) { return '· ' + m; }).join('\n')
      + '\n\n填報連結：' + FORM_URL;
    pushLineTo_(gid, msg, token);
  });
}

function missingForGroup_(g) {
  var tz = Session.getScriptTimeZone();
  var s = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd').split('-');
  var y = +s[0], m = +s[1], d = +s[2];
  var wk = isoWeekKey_(y, m, d), range = rangeOf_(y, m, d);
  var sh = getSheet_(), submitted = {}, last = sh.getLastRow();
  if (last > 1) {
    var vals = sh.getRange(2, 2, last - 1, 3).getValues(); // 週次,起訖,姓名
    vals.forEach(function (r) { if (String(r[0]) === wk) submitted[String(r[2])] = true; });
  }
  var missing = g.members.filter(function (mm) { return !submitted[mm]; });
  return { week: wk, range: range, missing: missing };
}

function pushLineTo_(gid, text, token) {
  token = token || PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to: gid, messages: [{ type: 'text', text: text }] }),
    muteHttpExceptions: true
  });
  Logger.log('push to ' + gid + ' status=' + res.getResponseCode() + ' ' + res.getContentText());
  return res.getResponseCode();
}

/* --- ISO 週次工具（伺服端）--- */
function isoWeekKey_(y, m, d) {
  var dt = new Date(Date.UTC(y, m - 1, d));
  var dayNum = dt.getUTCDay() || 7;
  dt.setUTCDate(dt.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(dt.getUTCFullYear(), 0, 1));
  var wk = Math.ceil((((dt - yearStart) / 86400000) + 1) / 7);
  return dt.getUTCFullYear() + '-W' + ('0' + wk).slice(-2);
}
function rangeOf_(y, m, d) {
  var dt = new Date(Date.UTC(y, m - 1, d));
  var day = (dt.getUTCDay() + 6) % 7;
  var mon = new Date(dt); mon.setUTCDate(dt.getUTCDate() - day);
  var sun = new Date(mon); sun.setUTCDate(mon.getUTCDate() + 6);
  return (mon.getUTCMonth() + 1) + '/' + mon.getUTCDate() + '–' + (sun.getUTCMonth() + 1) + '/' + sun.getUTCDate();
}

/* ============================ 一次性設定工具（在編輯器手動執行）============================ */

// 安裝／重設「每週五 16:30」自動催繳觸發器
function installReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendOverdueReminder') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendOverdueReminder')
    .timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(16).nearMinute(30).create();
  Logger.log('已安裝觸發器：每週五約 16:30 執行 sendOverdueReminder');
}

// 顯示各組已綁定的群組 ID + 最近捕捉到的 ID
function showGroupId() {
  var p = PropertiesService.getScriptProperties();
  GROUPS.forEach(function (g) { Logger.log(g.name + '（' + g.key + '）→ ' + p.getProperty('LINE_GROUP_ID_' + g.key)); });
  Logger.log('最近捕捉 LAST_GROUP_ID = ' + p.getProperty('LAST_GROUP_ID'));
}
function showLastGroupId() { Logger.log('LAST_GROUP_ID = ' + PropertiesService.getScriptProperties().getProperty('LAST_GROUP_ID')); }

// 把「最近在某群發話捕捉到的 LAST_GROUP_ID」綁定給第 N 組
function assignLineGroup1() { _assignLast(0); }
function assignLineGroup2() { _assignLast(1); }
function _assignLast(i) {
  var p = PropertiesService.getScriptProperties();
  if (!GROUPS[i]) { Logger.log('沒有第 ' + (i + 1) + ' 組'); return; }
  var last = p.getProperty('LAST_GROUP_ID');
  if (!last) { Logger.log('尚未捕捉到群組 ID，請先在該群發一句話'); return; }
  p.setProperty('LINE_GROUP_ID_' + GROUPS[i].key, last);
  Logger.log('已把最近捕捉到的群組 ID 綁定給【' + GROUPS[i].name + '】：' + last);
}

// 把舊的單一 LINE_GROUP_ID 搬到第 1 組（哈客組）— 升級時執行一次
function migrateFirstGroupId() {
  var p = PropertiesService.getScriptProperties();
  var old = p.getProperty('LINE_GROUP_ID');
  if (old) { p.setProperty('LINE_GROUP_ID_' + GROUPS[0].key, old); Logger.log('已把現有群組 ID 搬到【' + GROUPS[0].name + '】：' + old); }
  else Logger.log('找不到舊的 LINE_GROUP_ID（可能已搬過）');
}

// 對所有已綁定群組發一則測試訊息
function testLineHelloAll() {
  var p = PropertiesService.getScriptProperties(), token = p.getProperty('LINE_TOKEN');
  GROUPS.forEach(function (g) {
    var gid = p.getProperty('LINE_GROUP_ID_' + g.key);
    if (gid) pushLineTo_(gid, '【' + g.name + '】LINE 連線測試成功 ✅（可忽略）', token);
    else Logger.log(g.name + ' 尚未綁定群組 ID');
  });
}
function testReminderNow() { sendOverdueReminder(); }

/* ============================ 共用 ============================ */
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function ok_(s) { return ContentService.createTextOutput(s || 'OK'); }
