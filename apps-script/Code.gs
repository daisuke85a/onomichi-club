/**
 * おのみち部活さがし — Apps Script 側
 *
 * 使い方（README.md にも書いています）
 *  1. スプレッドシートを新規作成 → 拡張機能 → Apps Script
 *  2. このファイルの内容を Code.gs に、index.html を「HTML」ファイルとして追加
 *  3. 一度 setupSheet() を実行（見出しとサンプル行、入力規則が作られます）
 *  4. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *       実行ユーザー: 自分 ／ アクセスできるユーザー: 全員
 *  5. Google サイト → 挿入 → Apps Script → このデプロイを選ぶ
 */

const SHEET_NAME = 'クラブ一覧';

const HEADERS = ['公開', '種別', '種目', '名称', '参加対象', '対象学年', '費用', '活動日', '活動場所', '校区', '紹介', '連絡先', 'URL'];

function doGet(e) {
  // ?format=json で一覧を JSON で返す（Cloudflare Workers 版などの外部サイトから読む用）
  if (e && e.parameter && e.parameter.format === 'json') {
    return ContentService.createTextOutput(JSON.stringify(getClubs()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('おのみち部活さがし')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** ブラウザ側（index.html）から google.script.run.getClubs() で呼ばれる */
function getClubs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

  // getDisplayValues: 日付や金額を、シートに表示されている文字のまま受け取る
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(h => String(h).trim());

  return values.slice(1)
    .filter(row => row.some(c => String(c).trim() !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = String(row[i] || '').trim(); });
      return obj;
    })
    .filter(obj => String(obj['公開'] || '').toUpperCase() !== 'FALSE');
}

/** 初回だけ手動で実行：シート・見出し・入力規則・サンプル行を作る */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() > 0) {
    throw new Error('「' + SHEET_NAME + '」にすでにデータがあります。空にしてから実行してください。');
  }

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#DFECF2');
  sheet.setFrozenRows(1);

  const sample = [
    [true, 'クラブ', 'バスケットボール', '尾道シーサイド・バスケットクラブ U15', '男子・女子', '中1〜中3', '月3,000円', '週2日（火・木 19:00〜21:00）', '長江中学校 体育館', '長江中学校', '部活動の受け皿として2026年春に発足。経験者も初心者も歓迎。', '担当：山本（070-0000-0001）', 'https://example.com/seaside-bb'],
    [true, 'クラブ', '卓球', '尾道卓球クラブ', '男子・女子', '小5〜中3', '月2,000円', '週1日（土 9:00〜12:00）', '尾道市総合体育館 サブアリーナ', '久保中学校', 'ラケットは貸出あり。', '担当：佐藤', ''],
    [true, 'イベント', '体験会', 'バドミントン無料体験会', '男子・女子', '中1〜中3', '無料', '2026年10月25日（日）13:00〜15:00', '尾道市総合体育館', '全域', 'ラケット貸出あり。', '尾道市バドミントン協会', ''],
    [false, 'クラブ', 'サッカー', '（非公開のテスト行）', '男子', '中1〜中3', '月6,000円', '週3日', 'びんご運動公園', '高西中学校', '「公開」のチェックを外すとサイトに出ません。', '', ''],
  ];
  sheet.getRange(2, 1, sample.length, HEADERS.length).setValues(sample);

  const MAX = 500; // 入力規則をかけておく行数
  // 公開：チェックボックス
  sheet.getRange(2, 1, MAX, 1).insertCheckboxes();
  // 種別：プルダウン
  sheet.getRange(2, 2, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['クラブ', 'イベント'], true).setAllowInvalid(false).build());
  // 参加対象：プルダウン
  sheet.getRange(2, 5, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['男子・女子', '男子', '女子'], true).setAllowInvalid(true).build());
  // 校区：プルダウン（必要に応じて増減してください）
  const schools = ['全域', '長江中学校', '久保中学校', '栗原中学校', '日比崎中学校', '高西中学校', '向東中学校', '向島中学校', '因島南中学校', '瀬戸田中学校', '御調中学校'];
  sheet.getRange(2, 10, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(schools, true).setAllowInvalid(true).build());

  // 列幅
  const widths = [50, 80, 120, 260, 90, 90, 120, 220, 220, 120, 360, 200, 240];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.getRange(2, 11, MAX, 1).setWrap(true);
}
