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

// 掲載データのシート名。変えても動きますが（下の getDataSheet_ を参照）、
// 新しく作るときはこの名前になります
const SHEET_NAME = 'クラブ・イベント一覧';

// シート名が変わっていても、この3つの見出しがあれば掲載データのシートと判断します
const SHEET_KEY_HEADERS = ['公開', '種別', '名称'];

// チェックボックスや入力規則をあらかじめ用意しておく行数
const CHECKBOX_ROWS = 500;

const HEADERS = ['公開', '種別', '種目', '名称', '参加対象', '対象学年', '費用', '活動日', '活動場所', '校区', '紹介', '連絡先', 'URL', 'URLパスワード', '問い合わせ先メール',
  'アクセス', '駅・バス停・渡船場から徒歩10分以内', '駐輪場あり', '駐車場あり'];

const PASSWORD_COL = 'URLパスワード';
const CONTACT_COL = '問い合わせ先メール';

// 行き方の絞り込みに使う列。シートではチェックボックスで、TRUE なら該当
const ACCESS_FLAGS = ['駅・バス停・渡船場から徒歩10分以内', '駐輪場あり', '駐車場あり'];

// 無くても動く列。見張り（Health.gs）はこれらが欠けていても異常としません
const OPTIONAL_COLS = [PASSWORD_COL, CONTACT_COL, 'アクセス'].concat(ACCESS_FLAGS);

function doGet(e) {
  // ?health=1 … 自動の見張り（Health.gs）がデプロイ済みのページを外から確かめるための出口。
  // 件数だけを返し、団体の情報は出さない。
  if (e && e.parameter && e.parameter.health === '1') {
    const out = { ok: true, count: 0, at: new Date().toISOString() };
    try {
      out.count = getClubs().length;
    } catch (err) {
      out.ok = false;
      out.error = String((err && err.message) || err);
    }
    return ContentService.createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('おのみち部活さがし')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 掲載データのシートを返す。
 *
 * 名前で見つからないときは、1行目の見出しで見分ける。
 * 「1枚目のシート」を当てにすると、「引き継ぎ」など別のシートを
 * 読んでしまうため、当て推量はせず、見つからなければ止める。
 */
function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const byName = ss.getSheetByName(SHEET_NAME);
  if (byName) return byName;

  const found = ss.getSheets().filter(sh => {
    const lastCol = sh.getLastColumn();
    if (lastCol < SHEET_KEY_HEADERS.length) return false;
    const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h).trim());
    return SHEET_KEY_HEADERS.every(h => headers.indexOf(h) !== -1);
  });

  if (found.length === 1) return found[0];
  if (found.length > 1) {
    throw new Error('掲載データらしいシートが' + found.length + '枚あります（'
      + found.map(sh => sh.getName()).join('、') + '）。1枚だけにしてください。');
  }
  throw new Error('掲載データのシートが見つかりません。1行目に「'
    + SHEET_KEY_HEADERS.join('」「') + '」の見出しがあるシートを用意してください。');
}

/** シートを読み、1行を1オブジェクトにして返す（内部用。パスワードを含むので外に出さない） */
function readSheet_() {
  const sheet = getDataSheet_();

  // getDisplayValues: 日付や金額を、シートに表示されている文字のまま受け取る
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(h => String(h).trim());

  return values.slice(1)
    .map((row, i) => {
      const obj = { _row: i + 2 }; // シート上の実際の行番号（1行目は見出し）
      headers.forEach((h, j) => { obj[h] = String(row[j] || '').trim(); });
      return obj;
    })
    // 「名称」が入っている行だけを本物のデータとみなす。
    // チェックボックスを入れた列は空行でも 'FALSE' が入るため、
    // 「どれかの列が空でない」では空行を拾ってしまう
    .filter(obj => String(obj['名称'] || '').trim() !== '')
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
  // 秘密にすべき列。ブラウザには一切送らない
  const SECRET_COLS = [PASSWORD_COL, CONTACT_COL];

  return readSheet_().map((obj, i) => {
    const out = {};
    Object.keys(obj).forEach(k => { if (SECRET_COLS.indexOf(k) === -1) out[k] = obj[k]; });

    if (obj[PASSWORD_COL]) {
      out['URL'] = '';
      out._locked = true;
    }

    // 問い合わせ先が入っているかどうかだけを渡す。アドレスそのものは渡さない
    // （カードに出すとスパム収集の的になるため。送信は submitContact() がサーバー側で行う）
    out._canContact = isEmail_(obj[CONTACT_COL]);

    // 行き方の絞り込み用。シートのチェックボックスを真偽値にして渡す
    out._access = {};
    ACCESS_FLAGS.forEach(f => {
      out._access[f] = String(obj[f] || '').toUpperCase() === 'TRUE';
      delete out[f]; // 'TRUE'/'FALSE' の文字列は画面では使わない
    });

    // 試用中のA/B：1番目・3番目…はパターン1（問い合わせのみ）、
    //              2番目・4番目…はパターン2（問い合わせ＋申し込み）
    out._form = (i % 2 === 0) ? 1 : 2;

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

// 列を足したときの説明書き（右上に小さな印が付き、マウスを乗せると出ます）
const COLUMN_NOTES = {
  'URLパスワード': 'ここに合言葉を入れると、その行の URL は合言葉を入力した人にだけ表示されます。空欄なら誰でも見られます。',
  'アクセス': '最寄り駅やバス停からの行き方。例：JR尾道駅から徒歩10分／バス「長江口」下車すぐ。カードにそのまま表示されます。',
  '駅・バス停・渡船場から徒歩10分以内': '最寄りの駅・バス停・渡船場から歩いて10分以内ならチェック。送迎なしで通えるかの目安です。絞り込みに使われます。',
  '駐輪場あり': '自転車を停められるならチェック。絞り込みに使われます。',
  '駐車場あり': '送迎の車を停められるならチェック。絞り込みに使われます。',
  '問い合わせ先メール': 'サイトから届いた問い合わせ・申し込みの送り先。団体の了承を得たアドレスだけを入れてください。空欄なら、その団体のカードにフォームは出ません。',
};

const COLUMN_WIDTHS = { 'URLパスワード': 120, '問い合わせ先メール': 220, 'アクセス': 280,
  '駅・バス停・渡船場から徒歩10分以内': 250, '駐輪場あり': 90, '駐車場あり': 90 };

/**
 * 既に運用しているシートに、まだ無い列を右端に足す（手動で実行）。
 * HEADERS に項目が増えたら、これを1回実行すれば追いつきます。
 */
function addMissingColumns() {
  const sheet = getDataSheet_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(h => String(h).trim());

  const missing = HEADERS.filter(h => headers.indexOf(h) === -1);
  if (!missing.length) {
    throw new Error('足りない列はありません。');
  }

  let col = sheet.getLastColumn();
  missing.forEach(name => {
    col += 1;
    const cell = sheet.getRange(1, col);
    cell.setValue(name).setFontWeight('bold').setBackground('#DFECF2');
    sheet.setColumnWidth(col, COLUMN_WIDTHS[name] || 160);
    if (COLUMN_NOTES[name]) cell.setNote(COLUMN_NOTES[name]);
    // 行き方の列はチェックボックスにする
    if (ACCESS_FLAGS.indexOf(name) !== -1) {
      const rows = Math.max(Math.min(sheet.getMaxRows() - 1, CHECKBOX_ROWS), 1);
      sheet.getRange(2, col, rows, 1).insertCheckboxes();
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast('追加した列：' + missing.join('、'), 'おのみち部活さがし', 8);
}

/** 以前の名前。中身は addMissingColumns() と同じです */
function addPasswordColumn() {
  addMissingColumns();
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
    [true, 'クラブ', 'バスケットボール', '尾道シーサイド・バスケットクラブ U15', '男子・女子', '中1〜中3', '月3,000円', '週2日（火・木 19:00〜21:00）', '長江中学校 体育館', '長江中学校', '部活動の受け皿として2026年春に発足。経験者も初心者も歓迎。', '担当：山本（070-0000-0001）', 'https://example.com/seaside-bb', '', 'seaside-bb@example.com', 'JR尾道駅からバス10分「長江口」下車 徒歩3分', true, true, false],
    [true, 'クラブ', '卓球', '尾道卓球クラブ', '男子・女子', '小5〜中3', '月2,000円', '週1日（土 9:00〜12:00）', '尾道市総合体育館 サブアリーナ', '久保中学校', 'ラケットは貸出あり。', '担当：佐藤', 'https://example.com/ono-tt-form', 'onomichi2026', 'ono-tt@example.com', 'JR尾道駅から徒歩15分', true, true, true],
    [true, 'イベント', '体験会', 'バドミントン無料体験会', '男子・女子', '中1〜中3', '無料', '2026年10月25日（日）13:00〜15:00', '尾道市総合体育館', '全域', 'ラケット貸出あり。', '尾道市バドミントン協会', '', '', 'badminton@example.com', 'JR東尾道駅から徒歩20分／駐車場100台', false, true, true],
    [false, 'クラブ', 'サッカー', '（非公開のテスト行）', '男子', '中1〜中3', '月6,000円', '週3日', 'びんご運動公園', '高西中学校', '「公開」のチェックを外すとサイトに出ません。', '', '', '', '', '', false, false, false],
  ];
  sheet.getRange(2, 1, sample.length, HEADERS.length).setValues(sample);

  const MAX = CHECKBOX_ROWS; // 入力規則をかけておく行数
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
  const widths = [50, 80, 120, 260, 90, 90, 120, 220, 220, 120, 360, 200, 240, 120, 220, 280, 250, 90, 90];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.getRange(2, 11, MAX, 1).setWrap(true);

  // 行き方の列はチェックボックス
  ACCESS_FLAGS.forEach(name => {
    const i = HEADERS.indexOf(name);
    if (i !== -1) sheet.getRange(2, i + 1, MAX, 1).insertCheckboxes();
  });

  // 分かりにくい列に説明書きを付ける
  Object.keys(COLUMN_NOTES).forEach(name => {
    const i = HEADERS.indexOf(name);
    if (i !== -1) sheet.getRange(1, i + 1).setNote(COLUMN_NOTES[name]);
  });
}
