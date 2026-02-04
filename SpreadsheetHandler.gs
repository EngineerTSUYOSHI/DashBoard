/**
 * サイト別シートのデータをクリアする（ヘッダーは保持）
 */
function clearSiteData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("サイト別");
  
  // データが存在する最後の行を取得
  const lastRow = sheet.getLastRow();
  
  // 2行目以降にデータがある場合のみ実行
  if (lastRow > 1) {
    // A2から全範囲（最後の行・最後の列まで）を指定してクリア
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    console.log("サイト別シートのデータをクリアしました。");
  } else {
    console.log("クリアするデータはありません。");
  }
}

/**
 * 指定されたIDのスプレッドシートから本日分のデータを取得する
 * @param {string} spreadsheetId 
 * @param {string} targetSheetName 
 * @return {Array[]} 取得したデータの配列。本日分でなければnull
 */
function fetchTodayData(spreadsheetId, targetSheetName) {
  const targetSS = SpreadsheetApp.openById(spreadsheetId);
  const sheet = targetSS.getSheetByName(targetSheetName);
  
  if (!sheet) return null;

  // 1. B2セルから取得日を確認
  const checkDate = new Date(sheet.getRange("B2").getValue());
  const today = new Date();
  
  // 日付の比較（年・月・日が一致するか）
  const isToday = checkDate.getFullYear() === today.getFullYear() &&
                  checkDate.getMonth() === today.getMonth() &&
                  checkDate.getDate() === today.getDate();

  if (isToday) {
    // 2. 本日分であれば、ヘッダーを除いたデータ範囲を取得
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow < 2) return null; // データがない場合
    
    // A2から最終行・最終列まで取得
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    console.log(`✅ ${targetSS.getName()}: 本日分のデータ(${data.length}件)を取得しました。`);
    return data;
  } else {
    console.warn(`⚠️ ${targetSS.getName()}: 取得日が本日ではありません。スキップします。`);
    return null;
  }
}


function archiveSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("統合データ");
  const archiveSheet = ss.getSheetByName("集計");

  // 統合データのB2セル（合計件数）を確認
  const totalCount = summarySheet.getRange("B2").getValue();
  
  // データがない場合はスキップ（空、0、空文字列など）
  if (!totalCount || totalCount === 0 || totalCount === "") {
    console.log("⚠️ 統合シートにデータがないため、アーカイブ処理をスキップしました。");
    return;
  }

  // 1. 統合シートから最新結果（横一行）を取得
  // A列:日付, B列:合計, C列:中央値, D列以降:スキル名 ...
  const summaryHeaders = summarySheet.getRange(1, 4, 1, summarySheet.getLastColumn() - 3).getValues()[0];
  const summaryValues = summarySheet.getRange(2, 4, 1, summarySheet.getLastColumn() - 3).getValues()[0];
  const medianPrice = summarySheet.getRange("C2").getValue();

  // 2. 集計シートの既存ヘッダーを取得
  let archiveHeaders = archiveSheet.getRange(1, 1, 1, archiveSheet.getLastColumn()).getValues()[0];

  // 3. 日付を整形 (例: 26年2月1日(日))
  const now = new Date();
  const days = ["日", "月", "火", "水", "木", "金", "土"];
  const dateStr = Utilities.formatDate(now, "JST", "yy年M月d日") + "(" + days[now.getDay()] + ")";

  // 4. 新しい行の器を作成（初期値は空。A, B, C列は確定）
  let newRow = new Array(archiveHeaders.length).fill("");
  newRow[0] = dateStr;    // 収集日
  newRow[1] = totalCount; // 合計件数
  newRow[2] = medianPrice; // 中央値

  // 5. スキルを適切な列にマッピング
  summaryHeaders.forEach((skillName, index) => {
    if (!skillName) return;
    
    let colIndex = archiveHeaders.indexOf(skillName);
    
    // もし集計シートにないスキルなら、右端に追加
    if (colIndex === -1) {
      archiveSheet.getRange(1, archiveHeaders.length + 1).setValue(skillName);
      colIndex = archiveHeaders.length; // 新しい列のインデックスを取得
      archiveHeaders.push(skillName); // メモリ上のヘッダーリストも更新
    }
    
    // 既存または新規スキルの値をセット
    newRow[colIndex] = summaryValues[index];
  });

  // 6. 集計シートの末尾に1行追加
  archiveSheet.appendRow(newRow);
  console.log("✅ 集計シートへのアーカイブが完了しました。新スキルがあれば自動追加しました。");
}
