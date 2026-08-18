/**
 * 哈客組 週報填報系統 — Google Apps Script 後端
 * ------------------------------------------------------------
 *  1. doPost：
 *       (a) 網頁填報 → 依「週次+姓名」寫入/覆蓋試算表
 *       (b) LINE Webhook → 自動抓取群組 groupId（設定 LINE 用）
 *  2. doGet ：進度看板用 JSONP 讀取繳交狀態（不含內文）
 *  3. sendOverdueReminder：每週五 16:30 由觸發器自動執行，
 *       把「本週逾期未交」名單推播到 LINE 群組
 *
 *  ★ 改動此檔後，務必重新部署（管理部署作業→編輯→版本選「新版本」→部署），網址不變。
 */

/* ===== 需與 config.js 保持一致 ===== */
var TEAM_NAME = '哈客組';
var FORM_URL  = 'https://evaliuwrk-tech.github.io/hakka-weekly-report/';
var MEMBERS   = ['劉怡君（組長）', '涂政強', '黃敏琪', '劉美君'];   // 逾期名單依這份計算

var SHEET_NAME = 'Submissions';
var HEADERS = ['提交時間', '週次', '週次起訖', '姓名', '本週進度', '困難與處理', '下週規劃', '備註'];

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

  // (b) 這是 LINE Webhook 事件（含 events 陣列）→ 抓 groupId 存起來
  if (data && data.events) {
    try {
      var props = PropertiesService.getScriptProperties();
      data.events.forEach(function (ev) {
        var src = (ev && ev.source) || {};
        if (src.groupId) props.setProperty('LINE_GROUP_ID', src.groupId);
        else if (src.roomId) props.setProperty('LINE_GROUP_ID', src.roomId);
      });
    } catch (err2) {}
    return ContentService.createTextOutput('OK');   // LINE 需要回 200
  }

  // (a) 網頁填報 → upsert
  var lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    var sh = getSheet_();
    var row = [ new Date(), data.week || '', data.weekRange || '', data.name || '',
                data.progress || '', data.difficulties || '', data.nextPlan || '', data.note || '' ];
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
      var vals = sh.getRange(2, 1, last - 1, 4).getValues(); // 時間,週次,起訖,姓名
      rows = vals.map(function (r) {
        return {
          time: r[0] ? Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'MM/dd HH:mm') : '',
          week: r[1], weekRange: r[2], name: r[3]
        };
      });
    }
  } catch (err) { rows = []; }
  var out = JSON.stringify({ ok: true, rows: rows });
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) return ContentService.createTextOutput(cb + '(' + out + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

/* ============================ LINE 逾期催繳 ============================ */

// 每週五 16:30 由觸發器自動呼叫
function sendOverdueReminder() {
  var info = missingThisWeek_();
  if (!info.missing.length) return;               // 全員繳齊就不打擾
  var msg = '【' + TEAM_NAME + ' 週報逾期提醒】\n'
    + '本週（' + info.range + '）週報已過截止（週五 16:00），以下同仁尚未繳交，請儘快補上，謝謝！\n'
    + info.missing.map(function (m) { return '· ' + m; }).join('\n')
    + '\n\n填報連結：' + FORM_URL;
  pushLineGroup_(msg);
}

// 計算「本週」誰還沒交（用腳本時區判斷週次，與前端一致）
function missingThisWeek_() {
  var tz = Session.getScriptTimeZone();
  var s = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd').split('-');
  var y = +s[0], m = +s[1], d = +s[2];
  var wk = isoWeekKey_(y, m, d), range = rangeOf_(y, m, d);
  var sh = getSheet_(), submitted = {}, last = sh.getLastRow();
  if (last > 1) {
    var vals = sh.getRange(2, 2, last - 1, 3).getValues(); // 週次,起訖,姓名
    vals.forEach(function (r) { if (String(r[0]) === wk) submitted[String(r[2])] = true; });
  }
  var missing = MEMBERS.filter(function (mm) { return !submitted[mm]; });
  return { week: wk, range: range, missing: missing };
}

function pushLineGroup_(text) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('LINE_TOKEN');
  var gid = props.getProperty('LINE_GROUP_ID');
  if (!token) throw new Error('尚未設定 LINE_TOKEN（Channel access token）');
  if (!gid)  throw new Error('尚未抓到 LINE_GROUP_ID（請先把機器人加進群組並在群內發一句話）');
  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to: gid, messages: [{ type: 'text', text: text }] }),
    muteHttpExceptions: true
  });
  Logger.log('LINE push status=' + res.getResponseCode() + ' body=' + res.getContentText());
  return res.getResponseCode();
}

/* --- ISO 週次工具（伺服端，與前端同演算法）--- */
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

// ① 安裝／重設「每週五 16:30」自動催繳觸發器（執行一次即可）
function installReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendOverdueReminder') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendOverdueReminder')
    .timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(16).nearMinute(30).create();
  Logger.log('已安裝觸發器：每週五約 16:30 執行 sendOverdueReminder');
}

// ② 確認已抓到群組 ID（把機器人加進群、群內發一句話後，執行看看）
function showGroupId() {
  Logger.log('LINE_GROUP_ID = ' + PropertiesService.getScriptProperties().getProperty('LINE_GROUP_ID'));
}

// ③ 立即測試推播（設定好 Token 與群組後，手動執行一次驗證）
function testReminderNow() { sendOverdueReminder(); }

// ④ 直接發一則測試訊息到群組（不管有沒有人逾期）
function testLineHello() { pushLineGroup_('【' + TEAM_NAME + '】LINE 連線測試成功 ✅（可忽略）'); }

/* ============================ 共用 ============================ */
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function ok_(s) { return ContentService.createTextOutput(s || 'OK'); }
