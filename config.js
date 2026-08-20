/* ======================================================================
 *  ★★★  設定區：填報頁與看板頁共用，日後只要改這一個檔  ★★★
 *  多組共用：每一組一個 { key, name, members }。
 *    key    ：英數代號，用於看板網址 ?g=key 與 LINE 群組綁定（勿用中文/空格）
 *    name   ：組別顯示名稱
 *    members：該組成員（含組長），組長姓名後面加（組長）
 *  ★ 新增一組 = 在 groups 陣列複製一段、改 key/name/members；
 *    並到 Code.gs 的 GROUPS 也加同一組，然後重新部署。
 * ====================================================================== */
window.HAKKA_CONFIG = {
  systemName: "週報填報系統",
  appsScriptUrl: "https://script.google.com/macros/s/AKfycbxDA0YXWgKXwE-wz-KMHjbUkyTGbCnACCT55qrp9IcXln6vK7Mb0Q6PubdkUbOvkIM/exec",
  formUrl: "https://evaliuwrk-tech.github.io/hakka-weekly-report/",
  weeksToShow: 8,
  deadlineWeekday: 5,   // 1=週一 … 5=週五 … 7=週日
  deadlineHour: 16,     // 16 = 下午 4 點（週報截止時間）

  groups: [
    { key: "hakka", name: "哈客組",
      members: ["劉怡君（組長）", "涂政強", "黃敏琪", "劉美君"] },

    { key: "lsk",   name: "來上客組",
      members: ["邱嘉圓（組長）", "吳欣蓉（葉子）", "林郁婷", "蔡旻宏（阿蔡）"] }
  ]
};
