/**
 * のびのび飛竜くん ランキング（スプレッドシートに付けて使う GAS）
 *
 * 使い方
 *  1. スプレッドシートの「拡張機能 → Apps Script」にこのコードを貼る
 *  2. 関数 setup を1回実行する（シート「ランキング」「NGワード」ができる）
 *  3. 「デプロイ → 新しいデプロイ → ウェブアプリ」
 *     実行するユーザー：自分 ／ アクセスできるユーザー：全員
 *  4. 出てきた URL をゲーム側の RANK_URL に入れる
 *
 * 管理
 *  ・「表示」の列を × にすると、その行はランキングに出なくなる（行は消さない）
 *  ・「NGワード」シートに1行1つで言葉を書くと、その言葉を含む名前は登録できない
 */
const SHEET_RANK = 'ランキング';
const SHEET_NG = 'NGワード';
const MAX_M = 150;      // これより長い記録は受け付けない
const TOP_N = 20;       // 返す人数
const NAME_MAX = 8;     // 名前の文字数

function setup() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_RANK) || ss.insertSheet(SHEET_RANK);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['日時', '名前', '長さ(m)', '端末の番号', '表示']);
    sh.setFrozenRows(1);
  }
  const ng = ss.getSheetByName(SHEET_NG) || ss.insertSheet(SHEET_NG);
  if (ng.getLastRow() === 0) ng.appendRow(['使ってほしくない言葉（1行に1つ）']);
}

function doGet() {
  return out_(board_(null));
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const name = String(d.name || '').normalize('NFKC').trim();
    const m = Number(d.m);
    const id = String(d.id || '');

    if (!/^[a-z0-9]{8,32}$/.test(id)) throw new Error('登録できませんでした');
    const len = Array.from(name).length;
    if (len < 1 || len > NAME_MAX) throw new Error('名前は1〜' + NAME_MAX + '文字で入れてね');
    if (!/^[ぁ-んァ-ヶー一-龠々a-zA-Z0-9]+$/.test(name)) throw new Error('名前に使えない文字が入っています');
    if (isNg_(name)) throw new Error('その名前は使えません');
    if (!(m >= 3 && m <= MAX_M) || Math.round(m) !== m) throw new Error('登録できませんでした');

    const cache = CacheService.getScriptCache();
    if (cache.get('wait_' + id)) throw new Error('少し待ってからもう一度ためしてね');
    cache.put('wait_' + id, '1', 5);

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      SpreadsheetApp.getActive().getSheetByName(SHEET_RANK).appendRow([new Date(), name, m, id, '○']);
    } finally {
      lock.releaseLock();
    }
    cache.remove('board');
    return out_(Object.assign({ ok: true }, board_(id)));
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

/** ランキングを作る。id を渡すと、その端末の順位も返す。 */
function board_(id) {
  const cache = CacheService.getScriptCache();
  if (!id) {
    const hit = cache.get('board');
    if (hit) return JSON.parse(hit);
  }
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RANK);
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, 5).getValues() : [];
  const weekStart = weekStart_();
  const all = rank_(rows, 0, id);
  const week = rank_(rows, weekStart, id);
  const res = { all: all.top, week: week.top };
  if (id) {
    res.rankAll = all.me;
    res.rankWeek = week.me;
  } else {
    cache.put('board', JSON.stringify(res), 20);
  }
  return res;
}

/** 端末ごとのいちばん良い記録だけを残して、長い順に並べる。 */
function rank_(rows, since, id) {
  const best = {};
  rows.forEach(function (r) {
    const when = r[0] instanceof Date ? r[0].getTime() : 0;
    const m = Number(r[2]);
    const dev = String(r[3]);
    if (String(r[4]).trim() === '×' || !(m > 0) || when < since) return;
    const b = best[dev];
    if (!b || m > b.m || (m === b.m && when < b.t)) best[dev] = { n: String(r[1]), m: m, t: when, d: dev };
  });
  const list = Object.keys(best).map(function (k) { return best[k]; });
  list.sort(function (a, b) { return b.m - a.m || a.t - b.t; });
  let me = 0;
  if (id) for (let i = 0; i < list.length; i++) if (list[i].d === id) { me = i + 1; break; }
  return { top: list.slice(0, TOP_N).map(function (x) { return { n: x.n, m: x.m }; }), me: me };
}

/** 今週の月曜 0:00（日本時間）の時刻 */
function weekStart_() {
  const JST = 9 * 3600 * 1000;
  const j = new Date(Date.now() + JST);
  const back = (j.getUTCDay() + 6) % 7;
  return Date.UTC(j.getUTCFullYear(), j.getUTCMonth(), j.getUTCDate() - back) - JST;
}

function isNg_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NG);
  if (!sh || sh.getLastRow() < 2) return false;
  const low = name.toLowerCase();
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().some(function (r) {
    const w = String(r[0]).normalize('NFKC').trim().toLowerCase();
    return w && low.indexOf(w) !== -1;
  });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
