/**
 * 館内モニター（店舗サイトの booth.html）へ「予約が入っている時間」を渡すWebアプリ。
 *
 * なぜ必要か:
 *   公開カレンダーのICSフィードには CORS ヘッダが付いていないため、ブラウザのJSからは
 *   直接読めない（2026-08-02 実測）。GASを1枚挟んでJSONに変換して配る。
 *
 * 設計メモ:
 * - 予約の読み取りは **CalendarApp 直読み**（2026-08-07変更）。Holiday.gs 導入時に
 *   calendar スコープが承認済みになったため、イベントタイトル「電話ブース予約 (名前)」から
 *   予約者名が取れる。CalendarApp が失敗したら従来の公開ICS（名前なし）へフォールバック
 *   ＝モニターは死なない
 * - **名前はスタッフキー一致時だけ返す**（2026-08-07 本人裁定・スタッフキー方式）:
 *   スクリプトプロパティ BOARD_STAFF_KEY と ?staff= が一致した応答にのみ n（名前）を含める。
 *   booth.html は公開URLのため、無条件で名前を載せるとネット上の誰でも予約者名を見られてしまう。
 *   店内モニター（Fire Stick）のURLにだけ ?staff=キー を付けて運用する
 * - 応答は30秒キャッシュ（匿名向け・スタッフ向けの2系統。混ざると匿名に名前が漏れる）
 * - 失敗してもモニターが落ちないよう、エラー時は busy:[] と ok:false を返す（キャッシュしない）
 */

// 予約カレンダーの公開ICS URL（スクリプトプロパティ BOOTH_ICS_URL に設定。CalendarApp が失敗したときの予備経路）
var ICS_URL = PropertiesService.getScriptProperties().getProperty('BOOTH_ICS_URL');
// 反映確認用の版マーカー。/exec の応答に rev として載せる＝「pushしたコードが本番に出ているか」を
// ブラウザやcurlから1発で判定できる（デプロイがバージョン固定だとpushしても古い版が出続けるため）
var BOARD_CODE_REV = '2026-08-07a';
var BOARD_CACHE_KEY = 'booth_busy_v2';          // 匿名向け（名前なし）
var BOARD_CACHE_KEY_STAFF = 'booth_busy_v2s';   // スタッフ向け（名前あり）
var BOARD_STAFF_KEY_PROP = 'BOARD_STAFF_KEY';   // スクリプトプロパティのキー名
var BOARD_CACHE_SEC = 30; // モニター側は60秒ごとに取得＝反映は最悪90秒（2026-08-05: 60→30へ短縮）
var BOARD_RANGE_DAYS = 21; // 今日から何日先まで返すか（モニター表示は1週間・余裕を持たせる）

/** Webアプリ本体: {ok, updated, busy:[{s,e,n?}, ...]} を返す（n はスタッフキー一致時のみ） */
function doGet(e) {
  var staffKey = PropertiesService.getScriptProperties().getProperty(BOARD_STAFF_KEY_PROP);
  var isStaff = !!(staffKey && e && e.parameter && e.parameter.staff === staffKey);
  var cacheKey = isStaff ? BOARD_CACHE_KEY_STAFF : BOARD_CACHE_KEY;

  var cache = CacheService.getScriptCache();
  var hit = cache.get(cacheKey);
  if (hit) return asJson(hit);

  var body;
  try {
    body = JSON.stringify({ ok: true, rev: BOARD_CODE_REV, staff: isStaff, updated: jstStamp(new Date()), busy: fetchBusy(isStaff) });
    cache.put(cacheKey, body, BOARD_CACHE_SEC);
  } catch (err) {
    body = JSON.stringify({ ok: false, error: String(err), updated: jstStamp(new Date()), busy: [] });
  }
  return asJson(body);
}

function asJson(text) {
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 表示範囲内の予約済み時間帯を [{s,e,n?}] で返す。
 * 主経路 = CalendarApp（名前が取れる）／失敗時 = 公開ICS（名前なし・旧経路）。
 * フォールバックがあるので、万一 calendar スコープの承認が切れても掲示自体は生き続ける。
 */
function fetchBusy(withNames) {
  var now = new Date();
  var from = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var to = new Date(from.getTime() + BOARD_RANGE_DAYS * 86400000);
  try {
    return fetchBusyFromCalendar(from, to, withNames);
  } catch (err) {
    return fetchBusyFromIcs(from, to);
  }
}

/** CalendarApp で予約カレンダー（既定カレンダー）を直読みする */
function fetchBusyFromCalendar(from, to, withNames) {
  var events = CalendarApp.getDefaultCalendar().getEvents(from, to);
  var out = [];
  events.forEach(function (ev) {
    var item = { s: jstStamp(ev.getStartTime()), e: jstStamp(ev.getEndTime()) };
    if (withNames) {
      var n = pickBookerName(ev.getTitle());
      if (n) item.n = n;
    }
    out.push(item);
  });
  out.sort(function (a, b) { return a.s < b.s ? -1 : 1; });
  return out;
}

/**
 * イベントタイトルから表示用の名前を取り出す。
 * - 予約スケジュール経由のイベントは「電話ブース予約 (山田太郎)」形式（2026-08-03 通知メール件名で実測）
 *   → 末尾の括弧の中身が予約者名
 * - Holiday.gs の自動ブロックは「休業」とだけ出す（タイトル全文は掲示に長すぎる）
 * - それ以外の手動予定はタイトルをそのまま（スタッフが書いた用途メモがそのまま案内になる）
 */
function pickBookerName(title) {
  title = String(title || '').trim();
  if (!title) return null;
  if (title === '休業（自動ブロック）') return '休業'; // Holiday.gs BLOCK_TITLE と揃える
  var m = title.match(/\(([^()]+)\)\s*$/);
  if (m) return m[1].trim();
  return title.length > 20 ? title.slice(0, 20) + '…' : title;
}

/** 旧経路: 公開ICSを取得して [{s,e}] を返す（名前は含まれない） */
function fetchBusyFromIcs(from, to) {
  var res = UrlFetchApp.fetch(ICS_URL, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    throw new Error('ICS取得に失敗: HTTP ' + res.getResponseCode());
  }
  // ICSは75オクテットで行が折り返される。継続行（行頭が空白/タブ）を畳んでから読む
  var text = res.getContentText().replace(/\r?\n[ \t]/g, '');

  var out = [];
  var blocks = text.split('BEGIN:VEVENT');
  for (var i = 1; i < blocks.length; i++) {
    var b = blocks[i].split('END:VEVENT')[0];
    var s = pickIcsDate(b, 'DTSTART');
    var e = pickIcsDate(b, 'DTEND');
    if (!s || !e) continue;
    if (e <= from || s >= to) continue; // 表示範囲外は捨てる
    out.push({ s: jstStamp(s), e: jstStamp(e) });
  }
  out.sort(function (a, b2) { return a.s < b2.s ? -1 : 1; });
  return out;
}

/**
 * DTSTART / DTEND の行から Date を作る。
 * 3つの形に対応: 末尾Z=UTC ／ TZID付き=その地域時刻 ／ 日付のみ=終日予定
 */
function pickIcsDate(block, key) {
  var m = block.match(new RegExp('^' + key + '[^:\\r\\n]*:([0-9TZ]+)', 'm'));
  if (!m) return null;
  var v = m[1];
  if (/^\d{8}$/.test(v)) { // 20260806（終日）
    return new Date(+v.substr(0, 4), +v.substr(4, 2) - 1, +v.substr(6, 2));
  }
  var y = +v.substr(0, 4), mo = +v.substr(4, 2) - 1, d = +v.substr(6, 2);
  var h = +v.substr(9, 2), mi = +v.substr(11, 2), se = +v.substr(13, 2);
  if (/Z$/.test(v)) return new Date(Date.UTC(y, mo, d, h, mi, se));
  return new Date(y, mo, d, h, mi, se); // TZID付き＝スクリプトのタイムゾーン(Asia/Tokyo)で解釈
}

/** JSTの 'YYYY-MM-DDTHH:mm' 文字列にする */
function jstStamp(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm");
}

/** 動作確認用: Webアプリを叩かずに、実行ログへ返却JSON（匿名／スタッフ両方）を出す */
function testBoardFeed() {
  var cache = CacheService.getScriptCache();
  cache.remove(BOARD_CACHE_KEY);
  cache.remove(BOARD_CACHE_KEY_STAFF);
  console.log('匿名（名前なしのはず）: ' + doGet().getContent());
  var key = PropertiesService.getScriptProperties().getProperty(BOARD_STAFF_KEY_PROP);
  if (key) {
    console.log('スタッフ（名前ありのはず）: ' + doGet({ parameter: { staff: key } }).getContent());
  } else {
    console.log('⚠ スタッフキー未設定: プロジェクトの設定 → スクリプトプロパティに ' + BOARD_STAFF_KEY_PROP + ' を追加してください');
  }
}
