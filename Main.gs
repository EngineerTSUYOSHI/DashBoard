/**
 * Main.gs
 * メイン処理（オーケストレーション）を担当
 */
function runIntegration() {
  writeLog("開始", "処理を開始しました")
  // 1. まずは「サイト別」シートをクリア
  clearSiteData();
  
  // 2. 設定シートからデータ取得先のアクセス情報を取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("設定");
  const siteInfos = configSheet.getRange("A2:C4").getValues(); // A:サイト名, B:ID, C:シート名
  
  const destinationSheet = ss.getSheetByName("サイト別");

  // 3. 各サイトのデータを取得して「サイト別」シートにコピー
  let errorSites = [];
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
      writeLog("成功", `${siteName} から ${filteredData.length} 件を抽出してコピーしました。`);
    } else {
      errorSites.push(siteName);
      console.log(`❌ ${siteName} からデータを取得できませんでした。`);
      writeLog("失敗", `${siteName} からデータを取得できませんでした。`);
    }
  });
}

// 4. データを取得出来なかったサイトがある場合はエラーを通知
// ToDo: ここに通知ロジックを書く
if (errorSites.length > 0) {
  const errorMessage = "以下のサイトからデータを取得できませんでした: " + errorSites.join(", ");
  console.log(`❌ ${errorMessage}`);

// 5. 統合シートのデータを集計シートにコピー
archiveSummary();
writeLog("終了", "処理を終了しました")
}
