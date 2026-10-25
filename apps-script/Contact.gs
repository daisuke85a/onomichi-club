/**
 * おのみち部活さがし — サイトからの問い合わせ・申し込み
 *
 * 画面のフォームから google.script.run.submitContact() で呼ばれ、
 * その団体の「問い合わせ先メール」宛にメールを送り、「問い合わせ履歴」シートに記録します。
 *
 * 団体のメールアドレスはブラウザに一切渡していません（Code.gs の getClubs を参照）。
 * 宛先はここでシートから引き直すため、送信先を外から指定することはできません。
 */

const CONTACT_LOG_SHEET = '問い合わせ履歴';
const CONTACT_LOG_HEADERS = ['受付日時', '種別', '団体名', '行', 'お名前', '連絡先', '内容', '送信先', '結果'];

const CONTACT = {
  // 運用者にも控えを送る（試用中は true 推奨。届いているか確認できます）
  ccOperator: true,

  // 濫用対策。1時間あたりの上限
  maxPerClubPerHour: 5,
  maxTotalPerHour: 30,

  // フォームを開いてから送信までの最短秒数（これより速いものは機械とみなす）
  minFillSeconds: 3,

  // 入力の長さ上限
  maxNameLength: 100,
  maxContactLength: 200,
  maxMessageLength: 2000,
};

/**
 * フォームの送信を受け取る。
 *
 * @param {Object} p {row, type:'inquiry'|'apply', name, contact, message, hp, openedAt}
 * @return {{ok: boolean, message: string}}
 */
function submitContact(p) {
  p = p || {};

  // --- 罠フィールド。人には見えない欄に何か入っていたら機械 ---
  // 送信できたように見せて、実際には何もしない（相手に検知させないため）
  if (String(p.hp || '').trim() !== '') {
    return { ok: true, message: '送信しました。' };
  }

  // --- 速すぎる送信をはじく ---
  const elapsed = (Date.now() - Number(p.openedAt || 0)) / 1000;
  if (!(elapsed >= CONTACT.minFillSeconds)) {
    return { ok: false, message: '送信が早すぎます。少し時間をおいてお試しください。' };
  }

  // --- 入力の確認 ---
  const type = (p.type === 'apply') ? 'apply' : 'inquiry';
  const label = (type === 'apply') ? 'お申し込み' : 'お問い合わせ';

  const name = clean_(p.name, CONTACT.maxNameLength);
  const contact = clean_(p.contact, CONTACT.maxContactLength);
  const message = clean_(p.message, CONTACT.maxMessageLength);

  if (!name) return { ok: false, message: 'お名前を入力してください。' };
  if (!contact) return { ok: false, message: 'ご連絡先を入力してください。' };
  if (!message) return { ok: false, message: (type === 'apply' ? '申し込み内容' : 'お問い合わせ内容') + 'を入力してください。' };

  // --- 宛先はシートから引き直す（外から指定させない） ---
  const club = readSheet_().filter(o => o._row === Number(p.row))[0];
  if (!club) return { ok: false, message: 'この団体は現在受け付けていません。' };

  const to = String(club[CONTACT_COL] || '').trim();
  if (!isEmail_(to)) return { ok: false, message: 'この団体は現在受け付けていません。' };

  const clubName = club['名称'] || '（名称未設定）';

  // --- 送信回数の上限。重複送信も防ぐため鍵をかけて数える ---
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return { ok: false, message: '混み合っています。少し時間をおいてお試しください。' };
  }

  try {
    const over = overLimit_(club._row);
    if (over) return { ok: false, message: over };

    // --- 送信 ---
    const now = new Date();
    const stamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年M月d日 HH:mm');
    const body = [
      'おのみち部活さがし のページから、' + label + 'が届きました。',
      '',
      '■ 団体名',
      clubName,
      '',
      '■ お名前',
      name,
      '',
      '■ ご連絡先',
      contact,
      '',
      '■ ' + (type === 'apply' ? '申し込み内容' : 'お問い合わせ内容'),
      message,
      '',
      '■ 受付日時',
      stamp,
      '',
      '――――――',
      'このメールは、おのみち部活さがし のページから自動で送られています。',
      'お返事は、上のご連絡先に直接お願いします。',
    ].join('\n');

    const options = { name: 'おのみち部活さがし' };
    if (isEmail_(contact)) options.replyTo = contact; // 連絡先がメールならそのまま返信できるように
    if (CONTACT.ccOperator) {
      const operator = getNotifyAddresses_().to;
      if (operator.length) options.cc = operator.join(',');
    }

    let result = '送信';
    try {
      MailApp.sendEmail(to, '【おのみち部活さがし】' + label + '：' + clubName, body, options);
      bump_(club._row);
    } catch (err) {
      result = '失敗：' + String((err && err.message) || err);
    }

    logContact_([stamp, label, clubName, club._row, name, contact, message, to, result]);

    if (result !== '送信') {
      return { ok: false, message: '送信できませんでした。お手数ですが、カードに書かれた連絡先に直接ご連絡ください。' };
    }
    return {
      ok: true,
      message: (type === 'apply' ? 'お申し込みを送信しました。' : 'お問い合わせを送信しました。')
        + '団体から直接ご連絡があります。',
    };
  } finally {
    lock.releaseLock();
  }
}

/** 前後の空白を落とし、制御文字を外し、長すぎる場合は切る */
function clean_(v, max) {
  return String(v == null ? '' : v)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '')
    .trim()
    .slice(0, max);
}

/** 上限を超えていればその理由を返す。問題なければ空文字 */
function overLimit_(row) {
  const cache = CacheService.getScriptCache();
  const perClub = Number(cache.get('contact.row.' + row) || 0);
  const total = Number(cache.get('contact.all') || 0);

  if (perClub >= CONTACT.maxPerClubPerHour) {
    return 'この団体への送信が続いています。しばらく時間をおいてお試しください。';
  }
  if (total >= CONTACT.maxTotalPerHour) {
    return '現在混み合っています。しばらく時間をおいてお試しください。';
  }
  return '';
}

/** 送信回数を1つ増やす（1時間で消える） */
function bump_(row) {
  const cache = CacheService.getScriptCache();
  const a = 'contact.row.' + row;
  cache.put(a, String(Number(cache.get(a) || 0) + 1), 3600);
  cache.put('contact.all', String(Number(cache.get('contact.all') || 0) + 1), 3600);
}

/** 「問い合わせ履歴」シートに1行足す。シートが無ければ作る */
function logContact_(values) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONTACT_LOG_SHEET);
    if (!sheet) sheet = setupContactLog_();
    sheet.appendRow(values);
  } catch (err) {
    // 記録に失敗しても、送信そのものは止めない
    console.error('問い合わせ履歴に書けませんでした: ' + err);
  }
}

function setupContactLog_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONTACT_LOG_SHEET);
  if (sheet) return sheet;

  sheet = ss.insertSheet(CONTACT_LOG_SHEET);
  sheet.getRange(1, 1, 1, CONTACT_LOG_HEADERS.length).setValues([CONTACT_LOG_HEADERS])
    .setFontWeight('bold').setBackground('#DFECF2');
  sheet.setFrozenRows(1);
  [140, 90, 220, 50, 120, 200, 360, 200, 90].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.getRange(1, 1).setNote('サイトから届いた問い合わせ・申し込みの記録です。メールが届かなかった場合も、ここには残ります。');
  return sheet;
}

/** 手動で実行：「問い合わせ履歴」シートを用意する */
function setupContactLog() {
  setupContactLog_();
  SpreadsheetApp.getActiveSpreadsheet().toast('「' + CONTACT_LOG_SHEET + '」シートを用意しました。', 'おのみち部活さがし', 8);
}
