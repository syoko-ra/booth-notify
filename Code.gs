/**
 * 電話ブース予約 → Slack通知
 *
 * 予約用Googleアカウントの Gmail に届く予約メール（Googleカレンダー
 * 予約スケジュールの確定/キャンセル通知）を5分おきに検出し、Slackへ投稿する。
 *
 * 設置手順は README.md。必要なスクリプトプロパティ:
 *   SLACK_BOT_TOKEN   … botくんのBot User OAuth Token（xoxb-…）
 *   SLACK_CHANNEL_ID  … 投稿先チャンネルID（例 C0123456789）
 *
 * 設計メモ:
 * - 二重通知防止 = Gmailラベル「Slack通知済み」。Slack投稿が成功したスレッドにだけ付ける
 *   （投稿失敗時はラベルを付けず、次回実行で自動リトライ）
 * - メール本文のフォーマット変化に備え、パース失敗時は件名＋本文冒頭をそのまま流す（通知欠落ゼロ優先）
 */

var LABEL_NAME = 'Slack通知済み';
// 予約通知は「予約スケジュール名（電話ブース予約）」が件名に必ず入る。
// 送信元は calendar-notification ではなく **カレンダー所有者アカウント自身**（実測 2026-07-31）。
var SEARCH_QUERY = 'newer_than:2d -label:"' + LABEL_NAME +
  '" (subject:"電話ブース予約" OR from:calendar-notification@google.com)';

/**
 * テスト実行: 実予約なしで配線を一発検証する。
 * ①プロパティ2件の有無 ②Gmail検索が動くか ③Slack投稿が通るか（テスト通知1件が飛ぶ）
 * 結果はSlackと実行ログの両方に出る。何度実行してもラベル等の状態は変えない。
 */
function test() {
  var props = PropertiesService.getScriptProperties();
  var results = [];
  results.push((props.getProperty('SLACK_BOT_TOKEN') ? '✅' : '❌') + ' SLACK_BOT_TOKEN 設定');
  results.push((props.getProperty('SLACK_CHANNEL_ID') ? '✅' : '❌') + ' SLACK_CHANNEL_ID 設定');
  var mention = props.getProperty('SLACK_MENTION_USER_ID');
  results.push((mention ? '✅ メンション先: <@' + mention + '>' : 'ー メンションなし（SLACK_MENTION_USER_ID 未設定）'));

  var me = Session.getActiveUser().getEmail();
  results.push('ℹ 実行アカウント: ' + me + '（予約用アカウントで実行すること）');

  try {
    var recent = GmailApp.search('from:(calendar-notification@google.com) newer_than:7d', 0, 10).length;
    results.push('✅ Gmail検索OK（直近7日の予約系メール: ' + recent + '件）');
  } catch (e) {
    results.push('❌ Gmail検索エラー: ' + e);
  }

  var text = '🧪 *電話ブースSlack通知のテスト*\n' + results.join('\n') +
    '\nこのメッセージが見えていればSlack配線は正常です。';
  var ok = postToSlack(text);
  results.push(ok ? '✅ Slack投稿OK' : '❌ Slack投稿失敗（実行ログのエラーを確認: not_in_channel=bot未招待 / channel_not_found=ID間違い / invalid_auth=トークン間違い）');
  console.log(results.join('\n'));
}

/**
 * 診断: 「通知が来ない」ときに実行する。実物のメールを見て原因を切り分ける。
 * ①トリガーが生きているか ②現行クエリのヒット数 ③直近2日の受信メール一覧（送信元・件名・ラベル有無）
 * 結果はSlackと実行ログの両方へ。※件名と送信元だけを出し、本文は出さない
 */
function diagnose() {
  var out = ['🔍 *電話ブース通知の診断*'];

  var trigs = ScriptApp.getProjectTriggers().map(function (t) {
    return t.getHandlerFunction() + '(' + t.getEventType() + ')';
  });
  out.push(trigs.length ? '✅ トリガー: ' + trigs.join(', ') : '❌ トリガーなし → setup を実行すること');

  try {
    out.push('現行クエリのヒット: ' + GmailApp.search(SEARCH_QUERY, 0, 20).length + '件');
    out.push('（クエリ: `' + SEARCH_QUERY + '`）');
  } catch (e) {
    out.push('❌ クエリ実行エラー: ' + e);
  }

  var label = GmailApp.getUserLabelByName(LABEL_NAME);
  out.push(label ? '✅ ラベル「' + LABEL_NAME + '」あり（付与済み ' + label.getThreads(0, 50).length + '件）'
                 : 'ー ラベル未作成');

  out.push('--- 直近2日の受信メール（送信元 / 件名）---');
  var recent = GmailApp.search('newer_than:2d', 0, 15);
  if (!recent.length) {
    out.push('（受信メールが1件もありません）');
  } else {
    recent.forEach(function (th) {
      var m = th.getMessages()[th.getMessageCount() - 1];
      var labeled = th.getLabels().some(function (l) { return l.getName() === LABEL_NAME; });
      out.push((labeled ? '[済] ' : '[未] ') + m.getFrom() + ' / ' + m.getSubject());
    });
  }
  out.push('→ 予約通知の送信元がクエリと違う場合は、この送信元をClaudeに伝えてください');

  var text = out.join('\n');
  console.log(text);
  postToSlack(text);
}

/** 初回に1回だけ実行: 権限承認＋5分トリガー設置＋ラベル作成 */
function setup() {
  GmailApp.getUserLabelByName(LABEL_NAME) || GmailApp.createLabel(LABEL_NAME);
  // 自分のトリガーだけ張り替える（Holiday.gs の1時間トリガーを巻き込み削除しない）
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'checkBookings') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkBookings').timeBased().everyMinutes(5).create();
  checkBookings(); // その場で1回実行（動作確認を兼ねる）
}

/** 本体: 未通知の予約メールを探してSlackへ */
function checkBookings() {
  var label = GmailApp.getUserLabelByName(LABEL_NAME) || GmailApp.createLabel(LABEL_NAME);
  var threads = GmailApp.search(SEARCH_QUERY, 0, 20);
  threads.forEach(function (thread) {
    var msg = thread.getMessages()[thread.getMessageCount() - 1];
    var text = buildSlackText(msg.getSubject(), msg.getPlainBody());
    if (postToSlack(text)) {
      thread.addLabel(label);
    }
  });
}

/** 件名＋本文からSlack向けの短いメッセージを組み立てる */
function buildSlackText(subject, body) {
  var kind = '📩 予約関連の通知';
  if (/キャンセル|Cancell?ed/.test(subject)) kind = '🚫 *予約キャンセル*';
  else if (/更新|変更|Rescheduled|Updated/.test(subject)) kind = '🔁 *予約変更*';
  else if (/予約が完了|新しい予約|Booked|予約が確定/.test(subject)) kind = '📞 *新しい予約*';

  var lines = [kind];
  // 件名例: 予約が完了しました: 電話ブース予約 (山田太郎) - 2026年 8月 3日 (月) 午後5:30 ～ 午後6時 (JST) (…)
  var m = subject.match(/[:：]\s*.+?\s*\((.+?)\)\s*-\s*(.+?)\s*\(JST\)/);
  if (m) {
    lines.push('👤 ' + m[1]);
    lines.push('🕐 ' + m[2]);
  } else {
    lines.push('件名: ' + subject);
  }

  // 本文から要点をベストエフォートで抽出（日時・名前・メール・電話）
  var picked = [];
  var interesting = /(\d{1,2}月\s?\d{1,2}日|\d{1,2}:\d{2}|電話|メール|@|様|さん)/;
  var noise = /(https?:\/\/|カレンダー|Google|返信|通知|配信停止|unsubscribe|=+|-{4,})/i;
  body.split('\n').forEach(function (ln) {
    ln = ln.trim();
    if (ln && picked.length < 8 && interesting.test(ln) && !noise.test(ln)) picked.push(ln);
  });
  if (picked.length) {
    lines.push('---');
    lines = lines.concat(picked);
  }
  lines.push('（詳細は「電話ブース」カレンダー / 予約用アカウントのメール）');
  return lines.join('\n');
}

/**
 * Slack chat.postMessage。成功=true
 * SLACK_MENTION_USER_ID（例 U0XXXXXXXXX）が設定されていれば冒頭にメンションを付ける。
 * 相手を変える/やめる場合はスクリプトプロパティを書き換えるだけでよい（コード変更不要）。
 */
function postToSlack(text) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('SLACK_BOT_TOKEN');
  var channel = props.getProperty('SLACK_CHANNEL_ID');
  var mention = props.getProperty('SLACK_MENTION_USER_ID');
  if (mention) text = '<@' + mention + '>\n' + text;
  if (!token || !channel) {
    console.error('SLACK_BOT_TOKEN / SLACK_CHANNEL_ID が未設定');
    return false;
  }
  try {
    var res = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ channel: channel, text: text }),
      muteHttpExceptions: true,
    });
    var json = JSON.parse(res.getContentText());
    if (!json.ok) console.error('Slack API error: ' + json.error);
    return json.ok === true;
  } catch (e) {
    console.error('Slack送信失敗: ' + e);
    return false;
  }
}
