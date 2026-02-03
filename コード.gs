function testMultiSheetConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("設定");
  const TARGET_SHEET_NAME = "Python別、案件ランキング"; // ここに共通のシート名を入れてくれ
  
  // 設定シートのB2〜B4からIDを取得
  const sheetIds = configSheet.getRange("B2:B4").getValues().flat();
  
  sheetIds.forEach((id, index) => {
    if (!id) return; // 空セルはスキップ
    
    try {
      const targetSS = SpreadsheetApp.openById(id);
      const targetSheet = targetSS.getSheetByName(TARGET_SHEET_NAME);
      
      if (targetSheet) {
        console.log(`✅ サイト${index + 1} 接続成功: 「${targetSS.getName()}」内の「${TARGET_SHEET_NAME}」を認識しました。`);
      } else {
        console.warn(`⚠️ サイト${index + 1} 接続成功: ですが、シート「${TARGET_SHEET_NAME}」が見つかりません。`);
      }
    } catch (e) {
      console.error(`❌ サイト${index + 1} 接続失敗: ID "${id}" にアクセスできません。権限設定を確認してください。 エラー: ${e.message}`);
    }
  });
}

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


function runIntegration() {
  // 1. まずは「サイト別」シートをクリア
  clearSiteData();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("設定");
  const siteInfos = configSheet.getRange("A2:C4").getValues(); // A:サイト名, B:ID, C:シート名
  
  const destinationSheet = ss.getSheetByName("サイト別");

  siteInfos.forEach(info => {
    const [siteName, id, sheetName] = info;
    if (!id) return;
    
    // データ取得（fetchTodayDataは全列取ってくる前提）
    const rawData = fetchTodayData(id, "Python別、案件ランキング");
    
    if (rawData && rawData.length > 0) {
      // 2. 必要な列だけを抽出して並べ替え
      // インデックス: B=1, C=2, G=6, N=13
      const filteredData = rawData.map(row => {
        return [
          row[1],   // A列へ：元シートのB列（取得日？）
          siteName, // B列へ：スプシ名（設定シートのA列から取得）
          row[2],   // C列へ：元シートのC列（案件名？）
          row[6],   // D列へ：元シートのG列
          row[13]   // E列へ：元シートのN列
        ];
      });
      
      // 3. 「サイト別」シートの末尾に追記
      destinationSheet.getRange(
        destinationSheet.getLastRow() + 1, 
        1, 
        filteredData.length, 
        5 // A〜E列の5列固定
      ).setValues(filteredData);
      
      console.log(`✅ ${siteName} から ${filteredData.length} 件を抽出してコピーしました。`);
    }
  });
}

function archiveSummary() {
  writeLog("開始", "処理を開始しました")
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("統合データ");
  const archiveSheet = ss.getSheetByName("集計");

  // 1. 統合シートから最新結果（横一行）を取得
  // A列:日付, B列:合計, C列:中央値, D列以降:スキル名 ...
  const summaryHeaders = summarySheet.getRange(1, 4, 1, summarySheet.getLastColumn() - 3).getValues()[0];
  const summaryValues = summarySheet.getRange(2, 4, 1, summarySheet.getLastColumn() - 3).getValues()[0];
  const totalCount = summarySheet.getRange("B2").getValue();
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
      archiveHeaders.push(skillName); // メモリ上のヘッダーリストも更新
      newRow.push(summaryValues[index]); // 新しい列に値をセット
    } else {
      newRow[colIndex] = summaryValues[index]; // 既存の列に値をセット
    }
  });

  // 6. 集計シートの末尾に1行追加
  archiveSheet.appendRow(newRow);
  console.log("✅ 集計シートへのアーカイブが完了しました。新スキルがあれば自動追加しました。");
  writeLog("終了", "処理を終了しました")
}


/**
 * 実行ログを「実行ログ」シートの末尾に追記する
 * @param {string} status - 「成功」「失敗」「警告」など
 * @param {string} message - ログの内容
 */
function writeLog(status, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("実行ログ");

  // もし「実行ログ」という名前のシートがなければ、自動で作る（親切設計）
  if (!logSheet) {
    logSheet = ss.insertSheet("実行ログ");
    logSheet.appendRow(["実行日時", "ステータス", "内容"]); // ヘッダー作成
  }

  // 実行時の日時、ステータス、メッセージを1行追加
  const now = Utilities.formatDate(new Date(), "JST", "yyyy-MM-dd HH:mm:ss");
  logSheet.appendRow([now, status, message]);
}

