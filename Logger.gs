/**
 * 実行ログを「実行ログ」シートの末尾に追記する
 * @param {string} status - 「成功」「失敗」「警告」など
 * @param {string} message - ログの内容
 */
function writeLog(status, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("実行ログ");

  // もし「実行ログ」という名前のシートがなければ、自動で作る
  if (!logSheet) {
    logSheet = ss.insertSheet("実行ログ");
    logSheet.appendRow(["実行日時", "ステータス", "内容"]); // ヘッダー作成
  }

  // 実行時の日時、ステータス、メッセージを1行追加
  const now = Utilities.formatDate(new Date(), "JST", "yyyy-MM-dd HH:mm:ss");
  logSheet.appendRow([now, status, message]);
}
