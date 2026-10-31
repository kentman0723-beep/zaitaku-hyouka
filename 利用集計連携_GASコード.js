/**
 * knowbe作業時間データ → 利用集計スプレッドシート自動反映スクリプト (GAS)
 * 
 * 【設定手順】
 * 1. 利用集計スプレッドシートを開きます。
 *    URL: https://docs.google.com/spreadsheets/d/1i9Cn8Ay4qROo_WwZK4UCA-gGP0A-xLwSZ0PAtUA6DWE/edit?gid=1460989765#gid=1460989765
 * 2. 上部メニューの「拡張機能」 ➔ 「Apps Script」を開きます。
 * 3. 既存のコードを全て消去し、本ファイルの内容をそのまま貼り付けて保存（Ctrl+S）します。
 * 4. 右上の「デプロイ」 ➔ 「新しいデプロイ」をクリックします。
 * 5. 種類の歯車アイコン ➔ 「ウェブアプリ」を選択します。
 * 6. 以下の通り設定して「デプロイ」を実行します：
 *    - 説明: 利用集計連携
 *    - 次のユーザーとして実行: 自分 (Me)
 *    - アクセスできるユーザー: 全員 (Anyone)  ※認証なしでWebアプリから通信するために必須です
 * 7. 発行された「ウェブアプリのURL（https://script.google.com/macros/s/.../exec）」をコピーし、
 *    Webアプリ画面の「Google連携設定」欄に貼り付けて「設定を保存」します。
 */

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 対象シートの取得（GID指定があれば検索、なければアクティブシート）
    let sheet = null;
    if (data.sheetGid) {
      const sheets = ss.getSheets();
      sheet = sheets.find(s => s.getSheetId().toString() === data.sheetGid.toString());
    }
    if (!sheet) {
      sheet = ss.getActiveSheet();
    }

    // 氏名列 (F列 = 6列目) の行15〜最終行を取得
    const lastRow = sheet.getLastRow();
    if (lastRow < 15) {
      return responseJSON({ success: false, error: "利用者のデータ行（15行目以降）が見つかりません" });
    }

    const nameRange = sheet.getRange(15, 6, lastRow - 14, 1);
    const names = nameRange.getValues().map(row => (row[0] || "").toString().replace(/[\s　]/g, ""));

    // 日付1日〜31日の列範囲 (15列目/O列 〜 76列目/BX列: 計62列)
    // 1日: O(15) 予実, P(16) 食数
    // 2日: Q(17) 予実, R(18) 食数
    // ...
    // d日: (15 + (d-1)*2) 予実, (16 + (d-1)*2) 食数
    const totalCols = 31 * 2; // 62列
    const dataRange = sheet.getRange(15, 15, lastRow - 14, totalCols);
    const gridValues = dataRange.getValues();

    let updatedCount = 0;
    const notFoundNames = [];

    // 各利用者のデータをグリッドにマッピング
    data.users.forEach(user => {
      const targetName = (user.name || "").replace(/[\s　]/g, "");
      const rowIndex = names.findIndex(n => n === targetName);

      if (rowIndex === -1) {
        notFoundNames.push(user.name);
        return;
      }

      // 1日〜31日の実績を反映
      for (let day = 1; day <= 31; day++) {
        const colStatusIdx = (day - 1) * 2;     // 予実
        const colMealIdx = colStatusIdx + 1;     // 食数

        const dayRecord = user.records[day];
        if (dayRecord) {
          gridValues[rowIndex][colStatusIdx] = dayRecord.status || "";
          gridValues[rowIndex][colMealIdx] = dayRecord.meal || "";
        } else if (data.clearMissingDays) {
          // 作業時間データがない日はクリア（欠席扱い）
          gridValues[rowIndex][colStatusIdx] = "";
          gridValues[rowIndex][colMealIdx] = "";
        }
      }
      updatedCount++;
    });

    // スプレッドシートへ一括書き込み（高速）
    dataRange.setValues(gridValues);

    return responseJSON({
      success: true,
      updatedUsers: updatedCount,
      notFoundNames: notFoundNames
    });

  } catch (err) {
    return responseJSON({ success: false, error: err.toString() });
  }
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
