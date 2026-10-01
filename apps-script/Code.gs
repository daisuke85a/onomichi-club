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

const HEADERS = ['公開', '種別', '種目', '名称', '参加対象', '対象学年', '費用', '活動日', '活動場所', '校区', '紹介', '連絡先', 'URL', 'URLパスワード'];

const PASSWORD_COL = 'URLパスワード';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('おのみち部活さがし')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** シートを読み、1行を1オブジェクトにして返す（内部用。パスワードを含むので外に出さない） */
function readSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

  // getDisplayValues: 日付や金額を、シートに表示されている文字のまま受け取る
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(h => String(h).trim());

  return values.slice(1)
    .map((row, i) => {
      const obj = { _row: i + 2 }; // シート上の実際の行番号（1行目は見出し）
      headers.forEach((h, j) => { obj[h] = String(row[j] || '').trim(); });
      return obj;
    })
    .filter(obj => headers.some(h => obj[h] !== ''))
    .filter(obj => String(obj['公開'] || '').toUpperCase() !== 'FALSE');
}

/**
 * ブラウザ側（index.html）から google.script.run.getClubs() で呼ばれる。
 *
 * 「URLパスワード」が入っている行は、URL をブラウザに送らず `_locked: true` だけを渡す。
 * URL は getUrl() で正しいパスワードを受け取ったときに初めて返す。
 * 画面側で隠すだけでは開発者ツールから見えてしまうため、ここで落とすことが重要。
 */
function getClubs() {
  return readSheet_().map(obj => {
    const out = {};
    Object.keys(obj).forEach(k => { if (k !== PASSWORD_COL) out[k] = obj[k]; });
    if (obj[PASSWORD_COL]) {
      out['URL'] = '';
      out._locked = true;
    }
    return out;
  });
}

/**
 * ロックされた行の URL を、パスワードと引き換えに返す。
 *
 * @param {number} row シート上の行番号（getClubs が返した _row）
 * @param {string} password 利用者が入力したパスワード
 * @return {{ok: boolean, url?: string}}
 */
function getUrl(row, password) {
  const target = readSheet_().filter(o => o._row === Number(row))[0];
  if (!target) return { ok: false };

  const expected = String(target[PASSWORD_COL] || '');
  if (!expected) return { ok: true, url: target['URL'] || '' }; // パスワード未設定＝誰でも見られる

  if (String(password || '') !== expected) {
    Utilities.sleep(1000); // 総当たりを遅くするための待ち時間
    return { ok: false };
  }
  return { ok: true, url: target['URL'] || '' };
}

/**
 * 既に運用しているシートに「URLパスワード」列を足す（1回だけ手動で実行）。
 * setupSheet() を実行済みのシートで、あとからこの機能を使うとき用。
 */
function addPasswordColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(h => String(h).trim());

  if (headers.indexOf(PASSWORD_COL) !== -1) {
    throw new Error('「' + PASSWORD_COL + '」列はすでにあります。');
  }

  const col = headers.length + 1;
  sheet.getRange(1, col).setValue(PASSWORD_COL).setFontWeight('bold').setBackground('#DFECF2');
  sheet.setColumnWidth(col, 120);
  sheet.getRange(1, col).setNote('ここにパスワードを入れると、その行の URL はパスワードを入力した人にだけ表示されます。空欄なら誰でも見られます。');
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
    [true, 'クラブ', 'バスケットボール', '尾道シーサイド・バスケットクラブ U15', '男子・女子', '中1〜中3', '月3,000円', '週2日（火・木 19:00〜21:00）', '長江中学校 体育館', '長江中学校', '部活動の受け皿として2026年春に発足。経験者も初心者も歓迎。', '担当：山本（070-0000-0001）', 'https://example.com/seaside-bb', ''],
    [true, 'クラブ', '卓球', '尾道卓球クラブ', '男子・女子', '小5〜中3', '月2,000円', '週1日（土 9:00〜12:00）', '尾道市総合体育館 サブアリーナ', '久保中学校', 'ラケットは貸出あり。', '担当：佐藤', 'https://example.com/ono-tt-form', 'onomichi2026'],
    [true, 'イベント', '体験会', 'バドミントン無料体験会', '男子・女子', '中1〜中3', '無料', '2026年10月25日（日）13:00〜15:00', '尾道市総合体育館', '全域', 'ラケット貸出あり。', '尾道市バドミントン協会', '', ''],
    [false, 'クラブ', 'サッカー', '（非公開のテスト行）', '男子', '中1〜中3', '月6,000円', '週3日', 'びんご運動公園', '高西中学校', '「公開」のチェックを外すとサイトに出ません。', '', '', ''],
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
  const widths = [50, 80, 120, 260, 90, 90, 120, 220, 220, 120, 360, 200, 240, 120];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.getRange(2, 11, MAX, 1).setWrap(true);
}
