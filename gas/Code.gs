/**
 * TOEIC 金フレ 単語テスト - Google Apps Script
 *
 * 【セットアップ手順】
 * 1. Google スプレッドシートを新規作成
 * 2. 拡張機能 → Apps Script を開く
 * 3. このコードを貼り付けて保存（Ctrl+S）
 * 4. 「デプロイ」→「新しいデプロイ」をクリック
 * 5. 種類: ウェブアプリ
 *    実行ユーザー: 自分
 *    アクセス: 全員
 * 6. 「デプロイ」→ 表示されたURLをコピー
 * 7. アプリの config.js の GAS_URL に貼り付ける
 *
 * ※ コードを変更した場合は「新しいデプロイ」を作り直してください
 */

const SHEET_NAME = 'Results';

function doPost(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        '日時', '名前', 'セット番号', '正答数', '合計', '正答率(%)', '間違えた単語', '出題順'
      ]);
      // ヘッダー行を太字にする
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    }

    const data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date(),
      data.name,
      data.setNumber,
      data.score,
      data.total,
      data.accuracy,
      (data.wrongWords || []).join(' / '),
      data.order === 'random' ? 'ランダム' : '順番通り'
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet || sheet.getLastRow() <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const rows = sheet.getDataRange().getValues().slice(1); // skip header

    const records = rows.map(row => ({
      timestamp: row[0] ? row[0].toString() : '',
      name:      row[1],
      setNumber: row[2],
      score:     row[3],
      total:     row[4],
      accuracy:  row[5],
      wrongWords: row[6],
      order:     row[7]
    }));

    return ContentService
      .createTextOutput(JSON.stringify(records))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
