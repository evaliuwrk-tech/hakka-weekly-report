/* ======================================================================
 *  ★★★  設定區：填報頁與看板頁共用，日後只要改這一個檔  ★★★
 *  1. teamName      ：組別名稱
 *  2. members       ：組長 + 組員姓名（填報下拉、看板列都用這份）
 *  3. appsScriptUrl ：Apps Script 部署網址（.../exec）
 *  4. weeksToShow   ：看板顯示最近幾週
 *  5. formUrl       ：填報頁網址（用於看板「複製催繳訊息」）
 *  6. deadlineWeekday / deadlineHour：截止時間，用來判斷「逾期」（5=週五）
 * ====================================================================== */
window.HAKKA_CONFIG = {
  teamName: "哈客組",
  members: ["劉怡君（組長）", "涂政強", "黃敏琪", "劉美君"],
  appsScriptUrl: "https://script.google.com/macros/s/AKfycbxDA0YXWgKXwE-wz-KMHjbUkyTGbCnACCT55qrp9IcXln6vK7Mb0Q6PubdkUbOvkIM/exec",
  weeksToShow: 8,
  formUrl: "https://evaliuwrk-tech.github.io/hakka-weekly-report/",
  deadlineWeekday: 5,   // 1=週一 … 5=週五 … 7=週日
  deadlineHour: 16      // 16 = 下午 4 點（週報截止時間）
};
