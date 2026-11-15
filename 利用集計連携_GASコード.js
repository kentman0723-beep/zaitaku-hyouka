/**
 * 在宅評価と利用集計 — Googleスプレッドシート連携 & クラウドデータ同期スクリプト (GAS)
 * 
 * 【機能】
 * 1. knowbe作業時間データ → 利用集計スプレッドシート自動反映
 * 2. 別PC・複数デバイス間での「利用者情報・週評価・月1達成度評価・設定」クラウド自動同期
 *    （スプレッドシート内に「_アプリデータ保存」シートを自動作成してデータを保持します）
 * 
 * 【設定手順】
 * 1. 利用集計スプレッドシートを開きます。
 *    URL: https://docs.google.com/spreadsheets/d/1i9Cn8Ay4qROo_WwZK4UCA-gGP0A-xLwSZ0PAtUA6DWE/edit?gid=1460989765#gid=1460989765
 * 2. 上部メニューの「拡張機能」 ➔ 「Apps Script」を開きます。
 * 3. 既存のコードを全て消去し、本ファイルの内容をそのまま貼り付けて保存（Ctrl+S）します。
 * 4. 右上の「デプロイ」 ➔ 「新しいデプロイ」（または「デプロイを管理」➔ 鉛筆マーク ➔ バージョン「新バージョン」）を選択します。
 * 5. 種類「ウェブアプリ」を選択し、以下の通り設定して「デプロイ」を実行します：
 *    - 説明: 在宅評価・利用集計連携
 *    - 次のユーザーとして実行: 自分 (Me)
 *    - アクセスできるユーザー: 全員 (Anyone)  ※複数PCから認証なしで通信するために必須です
 * 6. 発行された「ウェブアプリのURL（https://script.google.com/macros/s/.../exec）」をコピーし、
 *    Webアプリの「設定」または「利用集計連携」画面に貼り付けて保存します。
 */

const DB_SHEET_NAME = "_アプリデータ保存";
const CHUNK_SIZE = 30000; // Googleスプレッドシートの1セル50,000文字制限を回避する分割サイズ

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "loadCloudData";
    const accountId = (e && e.parameter && e.parameter.accountId) ? e.parameter.accountId : "ogura_kajimachi";

    if (action === "loadCloudData") {
      return handleLoadCloudData(ss, accountId);
    }
    return responseJSON({ success: true, message: "在宅評価と利用集計 GAS連携サーバー稼働中" });
  } catch (err) {
    return responseJSON({ success: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- A. クラウドデータ同期（複数PC間での利用者情報・評価データの保存/読込） ---
    if (data.action === "saveCloudData") {
      return handleSaveCloudData(ss, data);
    }
    if (data.action === "loadCloudData") {
      return handleLoadCloudData(ss, data.accountId || "ogura_kajimachi");
    }

    // --- B. knowbe作業時間データ → 利用集計シートへの反映 ---
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
    const totalCols = 31 * 2; // 62列
    const dataRange = sheet.getRange(15, 15, lastRow - 14, totalCols);
    const gridValues = dataRange.getValues();

    let updatedCount = 0;
    const notFoundNames = [];

    // 各利用者のデータをグリッドにマッピング
    (data.users || []).forEach(user => {
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

        const dayRecord = user.records ? user.records[day] : null;
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

/**
 * アプリの全データ（利用者・週評価・月1達成度評価・設定）を「_アプリデータ保存」シートに保存
 */
function handleSaveCloudData(ss, data) {
  const accountId = (data.accountId || "ogura_kajimachi").toString();
  const payloadStr = JSON.stringify(data.payload || {});
  const chunks = [];
  for (let i = 0; i < payloadStr.length; i += CHUNK_SIZE) {
    chunks.push(payloadStr.substring(i, i + CHUNK_SIZE));
  }
  if (chunks.length === 0) chunks.push("{}");

  let dbSheet = ss.getSheetByName(DB_SHEET_NAME);
  if (!dbSheet) {
    dbSheet = ss.insertSheet(DB_SHEET_NAME);
    dbSheet.getRange(1, 1, 1, 3).setValues([["account_id", "updated_at", "data_chunks"]]);
  }

  const lastRow = dbSheet.getLastRow();
  let targetRow = -1;
  if (lastRow >= 2) {
    const ids = dbSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === accountId) {
        targetRow = i + 2;
        break;
      }
    }
  }
  if (targetRow === -1) {
    targetRow = Math.max(lastRow + 1, 2);
  } else {
    // 既存行をクリア
    const maxCols = Math.max(dbSheet.getLastColumn(), chunks.length + 2);
    dbSheet.getRange(targetRow, 1, 1, maxCols).clearContent();
  }

  const nowIso = new Date().toISOString();
  const rowData = [accountId, nowIso, ...chunks];
  dbSheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

  return responseJSON({
    success: true,
    accountId: accountId,
    updatedAt: nowIso
  });
}

/**
 * 「_アプリデータ保存」シートから指定アカウントのデータを読み込み
 */
function handleLoadCloudData(ss, accountId) {
  const targetId = (accountId || "ogura_kajimachi").toString();
  const dbSheet = ss.getSheetByName(DB_SHEET_NAME);
  if (!dbSheet) {
    return responseJSON({ success: true, exists: false, payload: null });
  }

  const lastRow = dbSheet.getLastRow();
  const lastCol = dbSheet.getLastColumn();
  if (lastRow < 2 || lastCol < 3) {
    return responseJSON({ success: true, exists: false, payload: null });
  }

  const allValues = dbSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (let i = 0; i < allValues.length; i++) {
    const row = allValues[i];
    if (String(row[0]) === targetId) {
      const updatedAt = row[1];
      const chunks = [];
      for (let c = 2; c < row.length; c++) {
        if (row[c] !== null && row[c] !== undefined && String(row[c]) !== "") {
          chunks.push(String(row[c]));
        }
      }
      const jsonStr = chunks.join("");
      const payload = jsonStr ? JSON.parse(jsonStr) : null;
      return responseJSON({
        success: true,
        exists: true,
        accountId: targetId,
        updatedAt: updatedAt,
        payload: payload
      });
    }
  }

  return responseJSON({ success: true, exists: false, payload: null });
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
