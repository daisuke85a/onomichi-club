/**
 * おのみち部活さがし — 自動の見張り（ヘルスチェック）
 *
 * 1日1回、ページとデータが正常かを自動で確かめ、異常があればメールで知らせます。
 * Google 側の仕様変更などでページが表示されなくなったとき、気づかずに放置するのを防ぐためのものです。
 *
 * 【使い方】Apps Script エディタで setupHealthCheck() を1回だけ実行してください。
 *          「通知先」シートができるので、そこに知らせ先のメールアドレスを書きます（複数可）。
 *          止めるときは removeHealthCheck() を実行します。
 *          メールが届くか試すには testHealthCheck() を実行します。
 */

// 知らせ先を書くシート。1行1アドレスで、何人でも登録できます
const NOTIFY_SHEET_NAME = '通知先';
const NOTIFY_HEADERS = ['有効', 'メールアドレス', '担当・備考'];

const HEALTH = {
  // 見張る公開ページ。デプロイし直して URL が変わったらここも直す
  webAppUrl: 'https://script.google.com/macros/s/AKfycbyMGMMw59yvpsAf01-SKy4_thHLWFPivV0xYk20bzsioV0xnSXXdyqlQ-nkbt-Py3YrPw/exec',

  // 追加で見張りたいページがあれば URL を足す（Cloudflare 版など）。不要なら空のまま
  extraUrls: [],

  // 予備の知らせ先。「通知先」シートが空のときだけ使う（空ならスクリプトの持ち主に送る）
  notifyTo: '',

  // 異常が続いている間、何時間おきに催促メールを送るか
  repeatHours: 24,

  // 正常でも週に1回「異常なし」のメールを送る曜日（0=日, 1=月 … 6=土）。送らないなら -1
  weeklyDigestDay: 1,
};

/** 毎日1回トリガーから呼ばれる本体 */
function healthCheck() {
  const problems = [];
  const notes = [];

  // --- 1. シートがあるか、見出しが揃っているか ---
  let rows = null;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    if (!sheet) throw new Error('シートが見つかりません');

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(h => String(h).trim());
    const required = HEADERS.filter(h => h !== PASSWORD_COL); // パスワード列は任意
    const missing = required.filter(h => headers.indexOf(h) === -1);
    if (missing.length) problems.push('シートの見出しが足りません：' + missing.join('、'));
    notes.push('シートの見出し：' + headers.length + '列');
  } catch (err) {
    problems.push('シートを読めません：' + err.message);
  }

  // --- 2. データを組み立てられるか ---
  try {
    rows = getClubs();
    if (!Array.isArray(rows)) throw new Error('getClubs() が配列を返しませんでした');
    if (rows.length === 0) problems.push('公開中のクラブ・イベントが0件です（「公開」のチェックが全部外れている可能性があります）');
    notes.push('公開中の件数：' + rows.length + '件');
  } catch (err) {
    problems.push('データを組み立てられません：' + err.message);
  }

  // --- 3. パスワード付きの行の URL が漏れていないか ---
  if (Array.isArray(rows)) {
    const leaked = rows.filter(r => r._locked && r['URL']);
    if (leaked.length) {
      problems.push('パスワードで保護しているはずの URL が外に出ています（' + leaked.length + '件）。すぐに連絡先に相談してください');
    }
  }

  // --- 4. 公開ページが実際に開けるか ---
  [HEALTH.webAppUrl].concat(HEALTH.extraUrls || []).filter(Boolean).forEach(url => {
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: true });
      const code = res.getResponseCode();
      if (code !== 200) {
        problems.push('ページが開けません（HTTP ' + code + '）： ' + url);
        return;
      }
      const body = res.getContentText();
      if (body.indexOf('おのみち部活さがし') === -1 || body.indexOf('id="list"') === -1) {
        problems.push('ページは開けましたが、中身がいつもと違います： ' + url);
        return;
      }
      notes.push('ページ応答：正常（' + url.slice(0, 60) + '…）');
    } catch (err) {
      problems.push('ページに接続できません：' + err.message + '（' + url + '）');
    }
  });

  // --- 5. 知らせ先が正しく登録されているか ---
  const notify = getNotifyAddresses_();
  if (!notify.to.length) {
    problems.push('知らせ先のメールアドレスが1件も登録されていません');
  } else {
    notes.push('知らせ先：' + notify.to.join('、') + '（' + notify.source + '）');
  }
  if (notify.invalid.length) {
    problems.push('「' + NOTIFY_SHEET_NAME + '」シートに、メールアドレスとして読めない行があります：' + notify.invalid.join('、'));
  }

  report_(problems, notes);
}

/** 結果に応じてメールを出す。正常が続いている間は送らない */
function report_(problems, notes) {
  const props = PropertiesService.getScriptProperties();
  const now = new Date();
  const was = props.getProperty('health.status') || 'ok';
  const lastSent = Number(props.getProperty('health.lastSent') || 0);
  const hoursSince = (now.getTime() - lastSent) / 3600000;

  if (problems.length) {
    const firstTime = was === 'ok';
    if (firstTime || hoursSince >= HEALTH.repeatHours) {
      send_(
        (firstTime ? '【要確認】' : '【継続中】') + 'おのみち部活さがし に問題が出ています',
        [
          (firstTime ? 'ページまたはデータに問題が見つかりました。' : '前回お知らせした問題がまだ続いています。'),
          '',
          '■ 見つかった問題',
          problems.map(p => '・' + p).join('\n'),
          '',
          '■ 確認した時刻',
          Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年M月d日 HH:mm'),
          '',
          '■ 公開ページ',
          HEALTH.webAppUrl,
          '',
          '■ スプレッドシート',
          SpreadsheetApp.getActiveSpreadsheet().getUrl(),
          '',
          '対処が分からないときは、引き継ぎ資料の「どうにもならなくなったら」に書かれた連絡先に相談してください。',
        ].join('\n')
      );
      props.setProperty('health.lastSent', String(now.getTime()));
    }
    props.setProperty('health.status', 'ng');
    return;
  }

  // --- 正常 ---
  if (was === 'ng') {
    send_('【復旧】おのみち部活さがし は正常に戻りました',
      ['先ほどお知らせした問題は解消しました。', '', '■ 状態', notes.map(n => '・' + n).join('\n'),
       '', '■ 確認した時刻', Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年M月d日 HH:mm')].join('\n'));
    props.setProperty('health.lastSent', String(now.getTime()));
  } else if (HEALTH.weeklyDigestDay >= 0 && now.getDay() === HEALTH.weeklyDigestDay) {
    const lastDigest = props.getProperty('health.lastDigest') || '';
    const today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
    if (lastDigest !== today) {
      send_('おのみち部活さがし 週次レポート（異常なし）',
        ['今週も正常に動いています。', '', '■ 状態', notes.map(n => '・' + n).join('\n'),
         '', '■ 公開ページ', HEALTH.webAppUrl,
         '', 'このメールが届かなくなったら、見張りの仕組み自体が止まっている可能性があります。',
         'Apps Script の「トリガー」画面を確認してください。'].join('\n'));
      props.setProperty('health.lastDigest', today);
    }
  }
  props.setProperty('health.status', 'ok');
}

function isEmail_(s) {
  return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(String(s).trim());
}

/**
 * 知らせ先を決める。
 *   1. 「通知先」シートの、有効チェックが入っている行すべて
 *   2. 無ければ HEALTH.notifyTo
 *   3. それも無ければスクリプトの持ち主
 * 誰にも届かない事態を避けるため、必ずどれかに落ちるようにしてある。
 */
function getNotifyAddresses_() {
  const out = { to: [], invalid: [], source: '' };

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOTIFY_SHEET_NAME);
    if (sheet && sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, NOTIFY_HEADERS.length).getDisplayValues()
        .forEach(r => {
          const addr = String(r[1] || '').trim();
          if (!addr) return;
          if (String(r[0] || '').toUpperCase() === 'FALSE') return; // チェックを外した行は送らない
          if (isEmail_(addr)) {
            if (out.to.indexOf(addr) === -1) out.to.push(addr);
          } else {
            out.invalid.push(addr);
          }
        });
      if (out.to.length) out.source = '「' + NOTIFY_SHEET_NAME + '」シート（' + out.to.length + '件）';
    }
  } catch (err) {
    // シートが読めなくても通知自体は止めない
  }

  if (!out.to.length && HEALTH.notifyTo && isEmail_(HEALTH.notifyTo)) {
    out.to = [HEALTH.notifyTo.trim()];
    out.source = 'Health.gs の HEALTH.notifyTo';
  }
  if (!out.to.length) {
    const owner = Session.getEffectiveUser().getEmail();
    if (owner) { out.to = [owner]; out.source = 'スクリプトの持ち主'; }
  }
  return out;
}

function send_(subject, body) {
  const n = getNotifyAddresses_();
  if (!n.to.length) throw new Error('送信先のメールアドレスが分かりません。「' + NOTIFY_SHEET_NAME + '」シートに登録してください。');

  const extra = n.invalid.length
    ? '\n\n※「' + NOTIFY_SHEET_NAME + '」シートに、メールアドレスとして読めない行がありました：' + n.invalid.join('、')
    : '';
  MailApp.sendEmail(n.to.join(','), subject, body + extra);
}

/**
 * 「通知先」シートを用意する（すでにあれば何もしない）。
 * setupHealthCheck() から自動で呼ばれるので、普段は直接実行しなくて構いません。
 */
function setupNotifySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(NOTIFY_SHEET_NAME)) return ss.getSheetByName(NOTIFY_SHEET_NAME);

  const sheet = ss.insertSheet(NOTIFY_SHEET_NAME);
  sheet.getRange(1, 1, 1, NOTIFY_HEADERS.length).setValues([NOTIFY_HEADERS])
    .setFontWeight('bold').setBackground('#DFECF2');
  sheet.setFrozenRows(1);

  const owner = Session.getEffectiveUser().getEmail();
  sheet.getRange(2, 1, 1, NOTIFY_HEADERS.length).setValues([[true, owner || '', '管理者']]);

  const MAX = 50;
  sheet.getRange(2, 1, MAX, 1).insertCheckboxes();
  sheet.getRange(2, 1, MAX, 1).setValue(false);
  sheet.getRange(2, 1).setValue(true);
  [60, 280, 200].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.getRange(1, 2).setNote('1行に1アドレス。何人でも登録できます。左のチェックを外すと、その人には送られません。');
  return sheet;
}

/** 1回だけ実行：毎日1回の見張りを仕掛ける */
function setupHealthCheck() {
  setupNotifySheet();
  removeHealthCheck();
  ScriptApp.newTrigger('healthCheck').timeBased().everyDays(1).atHour(8).create();
  const n = getNotifyAddresses_();
  send_('おのみち部活さがし の見張りを開始しました',
    ['毎日8時台にページとデータを自動で確認し、問題があればお知らせします。',
     '', '知らせ先：' + n.to.join('、') + '（' + n.source + '）',
     '知らせ先を増やすには、スプレッドシートの「' + NOTIFY_SHEET_NAME + '」シートに行を足してください。',
     '週次レポート：' + (HEALTH.weeklyDigestDay >= 0 ? '毎週' + '日月火水木金土'.charAt(HEALTH.weeklyDigestDay) + '曜日' : 'なし'),
     '', '止めるときは Apps Script エディタで removeHealthCheck() を実行してください。'].join('\n'));
}

/** 見張りを止める */
function removeHealthCheck() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'healthCheck')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/** 手動で1回確認し、結果を必ずメールで送る（動作確認用） */
function testHealthCheck() {
  PropertiesService.getScriptProperties().deleteProperty('health.status');
  PropertiesService.getScriptProperties().deleteProperty('health.lastSent');
  healthCheck();
  const status = PropertiesService.getScriptProperties().getProperty('health.status');
  if (status === 'ok') {
    send_('おのみち部活さがし テスト確認：異常なし',
      ['テスト実行の結果、問題は見つかりませんでした。', '', '異常時にはこのアドレスにお知らせが届きます。'].join('\n'));
  }
}
