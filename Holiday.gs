/**
 * 休業日の予約自動ブロック
 *
 * 営業日カレンダー（公開ICS）を1時間おきに読み、「営業日」イベントが無い平日を
 * 休業日とみなして、電話ブースカレンダーに 10:00〜18:00 のブロック予定を自動作成する。
 * 予約スケジュールは自分のカレンダーの予定と衝突する枠を出さないため、これだけで
 * 休業日の予約枠が消える（手動予定が枠を塞ぐことは 2026-07-13 のPhase 1検証で実証済み）。
 *
 * 発端 = 2026-08-05 の実事故: 臨時休業日（営業日イベント削除済み）に予約が3件入った。
 * 経営者の要望「ドロップイン休みの日＝予約不可にしたいね」への恒久対応。
 *
 * 設計メモ:
 * - 営業日の読み取りは公開ICSを UrlFetch（Board.gsと同じ経路・営業日カレンダーへの共有設定が不要）
 * - 書き込みは CalendarApp（要 calendar スコープ＝初回に setupHoliday を実行して再承認する）
 * - 自己修復: 営業日が復活した日は、自動ブロック（タイトル完全一致）だけを削除する
 * - 安全弁: ICSから営業日が1日も読めないときは何も作らない（解析失敗で全平日を塞がない）
 * - 休業日に既存予約が居たらSlackへ警告（イベントIDで重複警告を防止）
 * - 当日に急に休業が決まった場合、反映は最大1時間後（2026-08-05 本人了承済み）
 */

// 営業日カレンダーの公開ICS URL（スクリプトプロパティ OPEN_DAYS_ICS_URL に設定）
var EIGYOBI_ICS_URL = PropertiesService.getScriptProperties().getProperty('OPEN_DAYS_ICS_URL');
var BLOCK_TITLE = '休業（自動ブロック）';
var BLOCK_RANGE_DAYS = 16;   // 予約受付は14日先まで＋余裕2日
var BLOCK_START_HOUR = 10;   // 受付時間帯（予約スケジュール側の設定と揃える）
var BLOCK_END_HOUR = 18;
var HOLIDAY_ALERTED_KEY = 'holiday_alerted_ids';   // 警告済み予約ID（重複警告防止）
var HOLIDAY_ERROR_KEY = 'holiday_last_error_day';  // エラー通知は1日1回まで

/** 初回に1回だけ実行: calendarスコープの承認＋1時間トリガー設置＋その場で1回同期 */
function setupHoliday() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncHolidayBlocks') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncHolidayBlocks').timeBased().everyHours(1).create();
  syncHolidayBlocks();
}

/** 本体: 休業日にブロックを作り、営業日に戻った日のブロックを消す */
function syncHolidayBlocks() {
  var open;
  try {
    open = fetchOpenDays();
  } catch (e) {
    reportHolidayError('営業日ICSの取得/解析に失敗: ' + e);
    return;
  }
  if (!open.count) {
    reportHolidayError('営業日が1日も読めませんでした（カレンダー非公開化やICS仕様変更の可能性）。安全のためブロックは作りません');
    return;
  }

  var cal = CalendarApp.getDefaultCalendar();
  var today = startOfToday();
  var created = [], removed = [], warnings = [];

  for (var i = 0; i < BLOCK_RANGE_DAYS; i++) {
    var day = new Date(today.getTime() + i * 86400000);
    if (day.getDay() === 0 || day.getDay() === 6) continue; // 土日は予約スケジュール側で常時受付外
    var s = new Date(day); s.setHours(BLOCK_START_HOUR, 0, 0, 0);
    var e = new Date(day); e.setHours(BLOCK_END_HOUR, 0, 0, 0);
    var label = dayLabel(day);
    var events = cal.getEvents(s, e);
    var blocks = events.filter(function (ev) { return ev.getTitle() === BLOCK_TITLE; });
    var others = events.filter(function (ev) { return ev.getTitle() !== BLOCK_TITLE; });

    if (open.days[dayKey(day)]) {
      // 営業日: 残っている自動ブロックだけ消す（予約・手動予定には触らない）
      blocks.forEach(function (ev) { ev.deleteEvent(); });
      if (blocks.length) removed.push(label);
    } else {
      // 休業日: ブロックが無ければ作る
      if (!blocks.length) {
        cal.createEvent(BLOCK_TITLE, s, e);
        created.push(label);
      }
      // 既存の予約が居たら警告（1予約につき1回だけ）
      others.forEach(function (ev) {
        if (markAlerted(ev.getId())) {
          warnings.push(label + ' ' +
            Utilities.formatDate(ev.getStartTime(), 'Asia/Tokyo', 'HH:mm') + '〜' +
            Utilities.formatDate(ev.getEndTime(), 'Asia/Tokyo', 'HH:mm'));
        }
      });
    }
  }

  var lines = [];
  if (created.length) lines.push('🚧 休業日の予約枠を自動ブロックしました: ' + created.join('・'));
  if (removed.length) lines.push('♻️ 営業日に戻ったのでブロックを解除しました: ' + removed.join('・'));
  if (warnings.length) {
    lines.push('⚠️ *休業日に既に予約が入っています（要連絡）*: ' + warnings.join(' / '));
    lines.push('（詳細は「電話ブース」カレンダー / 予約用アカウントのメール）');
  }
  if (lines.length) {
    console.log(lines.join('\n'));
    postToSlack('📅 *休業日ブロック*\n' + lines.join('\n'));
  } else {
    console.log('変更なし（営業日 ' + open.count + '日 / 窓 ' + BLOCK_RANGE_DAYS + '日）');
  }
}

/**
 * 営業日カレンダーの公開ICSから「今日〜BLOCK_RANGE_DAYS日先」の営業日を集める。
 * 返り値: { days: {'yyyy-MM-dd': true, ...}, count: 窓内の営業日数 }
 * - SUMMARYに「営業日」を含むVEVENTだけを見る（他の予定が混ざっても誤読しない）
 * - 繰り返し（RRULE WEEKLY）は BYDAY・EXDATE・UNTIL・INTERVAL を解釈して展開する
 * - 単発イベント（休業→再追加の日など）はそのまま営業日に数える
 */
function fetchOpenDays() {
  var res = UrlFetchApp.fetch(EIGYOBI_ICS_URL, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
  var text = res.getContentText().replace(/\r?\n[ \t]/g, ''); // 75オクテット折り返しを畳む

  var today = startOfToday();
  var days = {};
  var blocks = text.split('BEGIN:VEVENT');
  for (var i = 1; i < blocks.length; i++) {
    var b = blocks[i].split('END:VEVENT')[0];
    if (!/^SUMMARY:.*営業日/m.test(b)) continue;
    if (/^STATUS:CANCELLED/m.test(b)) continue;
    var start = pickIcsDate(b, 'DTSTART'); // Board.gsの共通関数（Z/TZID/日付のみ対応）
    if (!start) continue;

    var rrule = (b.match(/^RRULE:(.+)$/m) || [])[1];
    if (!rrule) {
      days[dayKey(start)] = true;
      continue;
    }

    var rule = {};
    rrule.split(';').forEach(function (kv) {
      var p = kv.split('=');
      rule[p[0]] = p[1];
    });
    if (rule.FREQ !== 'WEEKLY') continue; // 週次以外は現運用に無い＝安全側（営業日に数えない）

    var dowMap = { SU: 0, MO: 1, TU: 2, WE: 3, TH: 4, FR: 5, SA: 6 };
    var byday = {};
    (rule.BYDAY || '').split(',').forEach(function (d) {
      if (dowMap[d] !== undefined) byday[dowMap[d]] = true;
    });
    var interval = +(rule.INTERVAL || 1);
    // UNTILはUTC表記（…Z）のことがあるが、日付部の比較で実用上足りる（受付は10-18時JST＝日跨ぎしない）
    var untilKey = rule.UNTIL ? icsDigitsToKey(rule.UNTIL) : null;

    var exdates = {};
    var exMatches = b.match(/^EXDATE[^:]*:(.+)$/mg) || [];
    exMatches.forEach(function (line) {
      line.replace(/^EXDATE[^:]*:/, '').split(',').forEach(function (v) {
        exdates[icsDigitsToKey(v)] = true;
      });
    });

    var startMid = new Date(start.getFullYear(), start.getMonth(), start.getDate());
    for (var j = 0; j < BLOCK_RANGE_DAYS; j++) {
      var day = new Date(today.getTime() + j * 86400000);
      if (!byday[day.getDay()]) continue;
      if (day < startMid) continue;
      var key = dayKey(day);
      if (untilKey && key > untilKey) continue;
      if (exdates[key]) continue;
      if (interval > 1 && weeksBetween(startMid, day) % interval !== 0) continue;
      days[key] = true;
    }
  }

  var count = 0;
  for (var j2 = 0; j2 < BLOCK_RANGE_DAYS; j2++) {
    if (days[dayKey(new Date(today.getTime() + j2 * 86400000))]) count++;
  }
  return { days: days, count: count };
}

/** 予約IDを警告済みとして記録。初出ならtrue（＝今回警告すべき） */
function markAlerted(id) {
  var props = PropertiesService.getScriptProperties();
  var ids = JSON.parse(props.getProperty(HOLIDAY_ALERTED_KEY) || '[]');
  if (ids.indexOf(id) !== -1) return false;
  ids.push(id);
  if (ids.length > 200) ids = ids.slice(-200);
  props.setProperty(HOLIDAY_ALERTED_KEY, JSON.stringify(ids));
  return true;
}

/** エラーは1日1回だけSlackへ（1時間トリガーで毎回鳴らさない）。ログには毎回出す */
function reportHolidayError(msg) {
  console.error(msg);
  var props = PropertiesService.getScriptProperties();
  var todayKey = dayKey(new Date());
  if (props.getProperty(HOLIDAY_ERROR_KEY) === todayKey) return;
  if (postToSlack('❌ *休業日ブロックのエラー*\n' + msg)) {
    props.setProperty(HOLIDAY_ERROR_KEY, todayKey);
  }
}

function startOfToday() {
  var now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

function dayKey(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function dayLabel(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'M/d(E)');
}

/** '20260805T100000' 等のICS日時値 → 'yyyy-MM-dd'（先頭8桁の日付部だけ使う） */
function icsDigitsToKey(v) {
  var m = String(v).match(/^(\d{4})(\d{2})(\d{2})/);
  return m ? m[1] + '-' + m[2] + '-' + m[3] : '';
}

/** WKST=SU前提の週差（INTERVAL判定用） */
function weeksBetween(a, b) {
  var aw = a.getTime() - a.getDay() * 86400000;
  var bw = b.getTime() - b.getDay() * 86400000;
  return Math.round((bw - aw) / (7 * 86400000));
}

/** 動作確認用: カレンダーに書き込まず、窓内の営業日/休業日の判定だけをログに出す */
function testHolidayParse() {
  var open = fetchOpenDays();
  var today = startOfToday();
  var out = ['営業日 ' + open.count + '日 / 窓 ' + BLOCK_RANGE_DAYS + '日'];
  for (var i = 0; i < BLOCK_RANGE_DAYS; i++) {
    var day = new Date(today.getTime() + i * 86400000);
    var kind = (day.getDay() === 0 || day.getDay() === 6) ? '土日（対象外）'
      : open.days[dayKey(day)] ? '営業日' : '★休業日→ブロック対象';
    out.push(dayLabel(day) + ' ' + kind);
  }
  console.log(out.join('\n'));
}
