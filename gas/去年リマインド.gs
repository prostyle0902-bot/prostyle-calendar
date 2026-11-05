// ===== 去年の同じ月にやっていた現場を思い出させる =====
// 去年の同じ月の「予測表」と board の案件を読んで、
// 今年まだカレンダーに入っていない現場を知らせる。
//
// これは「現場管理カレンダー」とは別の、新しい Apps Script に入れるファイルです。
// 入れ方・つなぎ方は REMIND_README.md を見てください。
//
// 動かすのに必要なもの（プロジェクトの設定 → スクリプト プロパティ）
//   BOARD_API_KEY    board の APIキー
//   BOARD_API_TOKEN  board の APIトークン
//   REMIND_TOKEN     合言葉（index.html の REMIND_TOKEN と同じ文字列）
// ※ APIキーはコードに書かないでください。このファイルは GitHub に公開されています。

// カレンダーの Apps Script（index.html の API と同じURL）
const CAL_API = 'https://script.google.com/macros/s/AKfycbwV_UNDYqwlGhMeoEQiasU348YX_UEKXys7E0TkZaOMfrcXtTx6-JIisyLVMCW2yj3g/exec';

// board の案件を何ページまで見るか（1ページ100件）
const BOARD_MAX_PAGES = 30;

// 予測表の中で読む欄の見出し
const H_DATE   = '日付';
const H_PLACE  = '現場名';
const H_CLIENT = '請求先お客様名';
const H_AMOUNT = '請負金額';
const H_OUTSRC = '外注先';

/* 同じ現場なのに、年によって書き方がちがうものを、ここで教える。
   前と後ろに付くもの（支店名など）は残したまま、会社の名前のところだけそろえる。
   例：['佐原信用金庫', '佐原信金'] と書いておくと
       「佐原信用金庫本店」と「佐原信金本店」は同じ現場、
       「佐原信用金庫本店」と「佐原信金笹川支店」は別の現場になる。
   いちばん左に、正しい（長いほう）の書き方を置いてください。 */
const SAME_SITE = [
  ['佐原信用金庫', '佐原信金'],
  ['佐原コミュニティーセンター', '佐原コミセン'],
  ['横河ソリューションサービス', '横河SS', '横河'],
  ['ネオマファーム', 'ネオマフォーム'],
];

/* 予測表の「区分」の欄（現場名の左どなり）が、ここに並べたものの行は読まない。
   鳥害対策は毎年くりかえす仕事ではないので、思い出さなくてよい。 */
const SKIP_KINDS = ['鳩', '鳥害', '鳥害対策'];

// 「合計」の行まで来たら、そこから下は読まない
const STOP_WORDS = ['合計', '外注費', '売上'];


/* =========================================================
   1. 最初に1回だけ：自動更新の設定
   ========================================================= */

// 4時間ごとに下ごしらえをしておく（画面を開いたときに待たされないように）
function setupRemindTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'refreshRemind')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshRemind').timeBased().everyHours(4).create();
  refreshRemind();
}

// 今月と来月のぶんを計算して、控えに残しておく
function refreshRemind() {
  const now = new Date();
  for (let diff = 0; diff <= 1; diff++) {
    const d = new Date(now.getFullYear(), now.getMonth() + diff, 1);
    const ym = `${d.getFullYear()}-${('0' + (d.getMonth() + 1)).slice(-2)}`;
    try {
      cachePut_(ym, buildRemind_(ym));
    } catch (e) {
      Logger.log(`${ym} の下ごしらえに失敗: ${e}`);
    }
  }
}


/* =========================================================
   2. 画面からの呼び出し口
   ========================================================= */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = p.callback || '';
  try {
    if (String(p.token || '') !== prop_('REMIND_TOKEN')) {
      return reply_({ status: 'err', message: '合言葉が違います' }, cb);
    }
    if (p.action === 'ping') {
      return reply_({ status: 'ok', message: '動いています' }, cb);
    }
    if (p.action === 'remind') {
      const ym = normYm_(p.ym) || thisYm_();
      // ふだんは控えを返す。fresh=1 のときだけ、その場で計算し直す
      let data = String(p.fresh || '') === '1' ? null : cacheGet_(ym);
      if (!data) {
        data = buildRemind_(ym);
        cachePut_(ym, data);
      }
      return reply_(data, cb);
    }
    return reply_({ status: 'err', message: '知らない操作です：' + p.action }, cb);
  } catch (err) {
    return reply_({ status: 'err', message: String(err && err.message || err) }, cb);
  }
}

function reply_(obj, cb) {
  const body = JSON.stringify(obj);
  if (cb) {
    return ContentService.createTextOutput(`${cb}(${body});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}


/* =========================================================
   3. 中身を作る
   ========================================================= */

/* ym は '2026-11' の形。その月について
     ・去年の同じ月（2025-11）にやっていた現場を、予測表と board から集める
     ・今年のカレンダーと見比べて、まだ入っていないものを選ぶ
   を行う。 */
function buildRemind_(ym) {
  const y = +ym.slice(0, 4), m = +ym.slice(5, 7);
  const ly = y - 1;

  const fromSheet = readForecast_(ly, m);
  const fromBoard = readBoard_(ly, m);

  // 現場名でまとめる（日常清掃のように何日も出てくるものを1件にする）
  const sites = {};
  fromSheet.rows.concat(fromBoard.rows).forEach(r => {
    const key = tidy_(r.place);
    if (!key) return;
    const s = sites[key] || (sites[key] = {
      place: r.place, client: r.client || '', amount: 0, days: 0,
      first: '', last: '', kind: r.kind || '', from: [],
    });
    s.amount += Number(r.amount) || 0;
    s.days += 1;
    if (r.date && (!s.first || r.date < s.first)) s.first = r.date;
    if (r.date && (!s.last || r.date > s.last)) s.last = r.date;
    if (!s.client && r.client) s.client = r.client;
    if (!s.kind && r.kind) s.kind = r.kind;
    if (s.from.indexOf(r.from) < 0) s.from.push(r.from);
  });

  // 今年のカレンダー
  const cal = readCalendar_();
  const thisMonth = placesIn_(cal.events, y, m);
  const around = placesIn_(cal.events, y, m, 1);   // 前後1か月ぶんも見る

  const todo = [], done = [], moved = [];
  Object.keys(sites).forEach(key => {
    const s = sites[key];
    if (matches_(key, thisMonth)) done.push(s);
    else if (matches_(key, around)) moved.push(s);
    else todo.push(s);
  });

  // 金額の大きいものから先に出す
  const bySize = (a, b) => (b.amount - a.amount) || (b.days - a.days);
  todo.sort(bySize);
  moved.sort(bySize);

  return {
    status: 'ok',
    ym: ym,
    lastYm: `${ly}-${('0' + m).slice(-2)}`,
    todo: todo,
    moved: moved,
    counts: {
      lastYear: Object.keys(sites).length,
      inThisMonth: done.length,
      inNearbyMonth: moved.length,
      notYet: todo.length,
    },
    sources: {
      forecast: fromSheet.ok ? `ok（${fromSheet.rows.length}件）` : fromSheet.reason,
      board: fromBoard.ok ? `ok（${fromBoard.rows.length}件）` : fromBoard.reason,
      calendar: cal.ok ? 'ok' : cal.reason,
    },
    computedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm'),
  };
}


/* =========================================================
   4. 予測表を読む
   ========================================================= */

/* 「2025年11月_予測表」を Drive から探して、
   日勤と外注の「現場名・請求先お客様名・請負金額」を取り出す。
   飲食事業部の欄（店舗名）は現場ではないので読まない。 */
function readForecast_(year, month) {
  const names = [`${year}年${month}月_予測表`, `${year}年${month}月予測表`];
  let file = null;
  for (const n of names) {
    const it = DriveApp.getFilesByName(n);
    if (it.hasNext()) { file = it.next(); break; }
  }
  if (!file) return { ok: false, reason: `「${names[0]}」が Drive に見つかりませんでした`, rows: [] };

  let ss;
  try { ss = SpreadsheetApp.openById(file.getId()); }
  catch (e) { return { ok: false, reason: `「${file.getName()}」を開けませんでした：${e}`, rows: [] }; }

  const sheets = ss.getSheets();
  const sheet = sheets.filter(s => s.getName().indexOf('粗利') >= 0)[0] || sheets[0];
  const values = sheet.getDataRange().getValues();

  // 見出しの行を探す（「日付」が入っている行）
  let hr = -1;
  for (let i = 0; i < Math.min(8, values.length); i++) {
    if (values[i].some(v => tidy_(v) === H_DATE)) { hr = i; break; }
  }
  if (hr < 0) return { ok: false, reason: `「${file.getName()}」に見出しの行が見つかりませんでした`, rows: [] };

  const head = values[hr].map(tidy_);
  const blocks = splitBlocks_(head);
  if (!blocks.length) return { ok: false, reason: `「${file.getName()}」に現場名の欄が見つかりませんでした`, rows: [] };

  const rows = [];
  blocks.forEach(b => {
    let curDate = '';
    for (let i = hr + 1; i < values.length; i++) {
      const row = values[i];
      const dateCell = row[b.date];
      const placeCell = tidy_(row[b.place]);

      // 「合計」などの行まで来たら、この欄はここで終わり
      if (isStop_(dateCell) || isStop_(placeCell)) break;

      const d = readDate_(dateCell, year);
      if (d) curDate = d.m === month ? `${year}-${('0' + d.m).slice(-2)}-${('0' + d.d).slice(-2)}` : '';

      if (!curDate || !placeCell) continue;

      const client0 = b.client >= 0 ? String(row[b.client] || '').trim() : '';
      const amount0 = readMoney_(b.amount >= 0 ? row[b.amount] : 0);
      // 「壁大18000、お掃除21000」「山ノ内不足分を請求する」のような
      // 書き置きの行は現場ではないので飛ばす
      if (!client0 && !amount0) continue;
      if (b.genre >= 0 && skipKind_(row[b.genre])) continue;
      if (/\d\s*円/.test(String(row[b.place]))) continue;

      rows.push({
        date: curDate,
        place: String(row[b.place]).trim(),
        client: client0,
        amount: amount0,
        kind: b.kind,
        from: '予測表',
      });
    }
  });

  return { ok: true, reason: '', rows: rows };
}

/* 見出しの行を「日付」の位置で区切って、現場名のある欄だけ拾う。
   予測表は〈日勤〉〈飲食事業部〉〈外注〉が横に並んでいるので、この形で分ける。 */
function splitBlocks_(head) {
  const starts = [];
  head.forEach((h, i) => { if (h === H_DATE) starts.push(i); });

  const blocks = [];
  starts.forEach((from, n) => {
    const to = n + 1 < starts.length ? starts[n + 1] - 1 : head.length - 1;
    const find = name => {
      for (let i = from; i <= to; i++) if (head[i] === name) return i;
      return -1;
    };
    const place = find(H_PLACE);
    if (place < 0) return;      // 飲食事業部（店舗名）はここで外れる
    // 日付・曜日と現場名のあいだにある、見出しのない欄が「区分」（ビルメン・鳩など）
    let genre = -1;
    for (let i = from + 1; i < place; i++) if (!head[i]) { genre = i; break; }
    blocks.push({
      date: from,
      place: place,
      genre: genre,
      client: find(H_CLIENT),
      amount: find(H_AMOUNT),
      kind: find(H_OUTSRC) >= 0 ? '外注' : '日勤',
    });
  });
  return blocks;
}


/* =========================================================
   5. board を読む
   ========================================================= */

// 案件の日付が入っていそうな欄を、上から順に見る
const BOARD_DATE_KEYS = [
  'scheduled_on', 'scheduled_date', 'work_date', 'delivery_date', 'delivered_on',
  'order_date', 'ordered_on', 'issue_date', 'billing_date', 'billed_on',
  'start_date', 'date', 'closing_date', 'created_at', 'updated_at',
];

function readBoard_(year, month) {
  const key = prop_('BOARD_API_KEY'), tok = prop_('BOARD_API_TOKEN');
  if (!key || !tok) {
    // キーを入れるまでは board は使わない。予測表だけで出す
    return { ok: false, off: true, reason: '使っていません（APIキーを入れると使えます）', rows: [] };
  }

  const rows = [];
  let scanned = 0;
  for (let page = 1; page <= BOARD_MAX_PAGES; page++) {
    let res;
    try {
      res = UrlFetchApp.fetch(`https://api.the-board.jp/v1/projects?per_page=100&page=${page}`, {
        method: 'get',
        muteHttpExceptions: true,
        headers: { 'x-api-key': key, 'Authorization': 'Bearer ' + tok },
      });
    } catch (e) {
      return { ok: false, reason: `board につながりませんでした：${e}`, rows: rows };
    }
    if (res.getResponseCode() !== 200) {
      return {
        ok: false,
        reason: `board が ${res.getResponseCode()} を返しました：${String(res.getContentText()).slice(0, 200)}`,
        rows: rows,
      };
    }

    let body;
    try { body = JSON.parse(res.getContentText()); }
    catch (e) { return { ok: false, reason: `board の返事を読めませんでした：${e}`, rows: rows }; }

    const list = Array.isArray(body) ? body : (body.data || body.projects || body.items || []);
    if (!list.length) break;
    scanned += list.length;

    list.forEach(o => {
      const d = boardDate_(o);
      if (!d || d.getFullYear() !== year || d.getMonth() + 1 !== month) return;
      const place = String(o.name || o.subject || o.title || '').trim();
      if (!place) return;
      rows.push({
        date: Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd'),
        place: place,
        client: String((o.client && o.client.name) || o.client_name || '').trim(),
        amount: readMoney_(o.total != null ? o.total
          : o.total_amount != null ? o.total_amount
          : o.amount != null ? o.amount : 0),
        kind: '',
        from: 'board',
      });
    });

    const pg = body.pagination;
    if (pg && pg.has_more === false) break;
    if (!pg && list.length < 100) break;
  }

  if (!rows.length) {
    return { ok: false, reason: `board の案件 ${scanned} 件を見ましたが、${year}年${month}月のものはありませんでした`, rows: [] };
  }
  return { ok: true, reason: '', rows: rows };
}

function boardDate_(o) {
  for (const k of BOARD_DATE_KEYS) {
    const d = toDate_(o[k]);
    if (d) return d;
  }
  // 上の並びに無い場合は、「〜_date」「〜_on」で終わる欄を順に見る
  for (const k of Object.keys(o)) {
    if (!/(_date|_on)$/.test(k)) continue;
    const d = toDate_(o[k]);
    if (d) return d;
  }
  return null;
}

/* board のつなぎ具合をためす。
   実行して「実行ログ」を見ると、案件の欄の名前と中身が分かる。
   日付や金額の欄の名前がちがっていたら、このログを見て
   BOARD_DATE_KEYS や readBoard_ の中を直してください。 */
function boardTest() {
  const key = prop_('BOARD_API_KEY'), tok = prop_('BOARD_API_TOKEN');
  if (!key || !tok) { Logger.log('APIキー・APIトークンがまだ入っていません'); return; }
  const res = UrlFetchApp.fetch('https://api.the-board.jp/v1/projects?per_page=1&page=1', {
    method: 'get',
    muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'Authorization': 'Bearer ' + tok },
  });
  Logger.log(`返事の番号: ${res.getResponseCode()}`);
  Logger.log(String(res.getContentText()).slice(0, 3000));
}


/* =========================================================
   6. 今年のカレンダーを読む
   ========================================================= */

function readCalendar_() {
  let res;
  try { res = UrlFetchApp.fetch(`${CAL_API}?action=getAll`, { muteHttpExceptions: true, followRedirects: true }); }
  catch (e) { return { ok: false, reason: `カレンダーにつながりませんでした：${e}`, events: {} }; }
  if (res.getResponseCode() !== 200) {
    return { ok: false, reason: `カレンダーが ${res.getResponseCode()} を返しました`, events: {} };
  }

  const text = String(res.getContentText());
  let obj = null;
  try { obj = JSON.parse(text); }
  catch (e) {
    // cb({...}); の形で返ってきたときは、かっこの中だけ取り出す
    const i = text.indexOf('('), j = text.lastIndexOf(')');
    if (i >= 0 && j > i) { try { obj = JSON.parse(text.slice(i + 1, j)); } catch (e2) {} }
  }
  if (!obj) return { ok: false, reason: 'カレンダーの返事を読めませんでした', events: {} };

  const events = obj.events || obj;
  return { ok: true, reason: '', events: (events && typeof events === 'object') ? events : {} };
}

/* その月（slack か月ぶん前後に広げても可）のカレンダーに入っている現場名を集める */
function placesIn_(events, y, m, slack) {
  const pad = slack || 0;
  const from = new Date(y, m - 1 - pad, 1);
  const to = new Date(y, m + pad, 0);
  const out = [];
  Object.keys(events || {}).forEach(k => {
    const p = String(k).split('-');
    if (p.length < 3) return;
    const d = new Date(+p[0], +p[1] - 1, +p[2]);
    if (d < from || d > to) return;
    (events[k] || []).forEach(ev => {
      const t = tidy_(ev && ev.place);
      if (t) out.push(t);
    });
  });
  return out;
}

/* 現場名が合うかを見る。
   そろえた名前がぴったり同じか、どちらかがもう一方を含んでいれば「同じ現場」とみなす
   （「済生会宇都宮病院」と「済生会宇都宮病院(ガラス)」のような書き方のちがいを拾うため）。 */
function matches_(key0, list0) {
  const key = alias_(key0);
  const list = list0.map(alias_);
  for (const p of list) {
    if (p === key) return true;
    if (key.length >= 4 && p.indexOf(key) >= 0) return true;
    if (p.length >= 4 && key.indexOf(p) >= 0) return true;
    // かなの打ちまちがいも同じ現場とみなす（「コミュニティー」と「コミニティー」）。
    // ただし漢字がちがえば別の現場なので（「梅屋敷院」と「梅屋敷分院」は別の院）、
    // 漢字の部分がそっくり同じで、かなのちがいだけのときに限る。
    if (key.length >= 5 && p.length >= 5
        && kanaless_(key) === kanaless_(p)
        && dice_(key, p) >= 0.7) return true;
  }
  return false;
}

/* SAME_SITE に書いた言いかえを当てはめる。
   会社の名前のところだけ「正しい書き方」に直し、支店名などはそのまま残す */
function alias_(key) {
  for (const group of SAME_SITE) {
    const head = tidy_(group[0]);
    for (const w of group) {
      const t = tidy_(w);
      if (t && key.indexOf(t) >= 0) return head + key.replace(t, '');
    }
  }
  return key;
}

/* かなを取りのぞいた、漢字・数字・英字だけの形。
   tidy_ を通したあとはカタカナもひらがなになっているので、これで漢字だけが残る */
function kanaless_(s) {
  return String(s).replace(/[ぁ-ゖ]/g, '');
}

/* ふたつの名前の似ぐあいを 0〜1 で返す（2文字ずつに切って、重なりを数える） */
function dice_(a, b) {
  if (a.length < 2 || b.length < 2) return a === b ? 1 : 0;
  const grams = t => {
    const m = {};
    for (let i = 0; i < t.length - 1; i++) {
      const k = t.substr(i, 2);
      m[k] = (m[k] || 0) + 1;
    }
    return m;
  };
  const A = grams(a), B = grams(b);
  let common = 0, na = 0, nb = 0;
  Object.keys(A).forEach(k => { na += A[k]; if (B[k]) common += Math.min(A[k], B[k]); });
  Object.keys(B).forEach(k => { nb += B[k]; });
  return na + nb ? (2 * common) / (na + nb) : 0;
}


/* =========================================================
   7. 小さな道具
   ========================================================= */

/* 見出しや名前を見比べるためにそろえる。
   空白・記号を取り、全角を半角に、カタカナをひらがなにして、
   書き方のゆれやすい「の・ノ・之」は取ってしまう
   （「山ノ内病院」と「山之内病院」、「アジカン」と「あじかん」を同じ扱いにするため）。 */
function tidy_(v) {
  let s = String(v == null ? '' : v);
  s = s.replace(/[Ａ-Ｚａ-ｚ０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  s = s.replace(/[（]/g, '(').replace(/[）]/g, ')');
  s = s.replace(/[\s　・,，、.。\-ー－_／\/]/g, '');
  s = s.replace(/[ァ-ヶ]/g, c => String.fromCharCode(c.charCodeAt(0) - 0x60));
  s = s.replace(/[のノ之]/g, '');
  return s.toLowerCase();
}

// 読まない区分（鳥害対策など）かどうか
function skipKind_(v) {
  const t = tidy_(v);
  if (!t) return false;
  return SKIP_KINDS.some(w => t.indexOf(tidy_(w)) >= 0);
}

function isStop_(v) {
  const s = tidy_(v);
  if (!s) return false;
  return STOP_WORDS.some(w => s.indexOf(tidy_(w)) >= 0);
}

// 「11/01」「11月1日」や、日付そのものを読む
function readDate_(v, year) {
  if (v instanceof Date) return { m: v.getMonth() + 1, d: v.getDate() };
  const s = String(v == null ? '' : v).replace(/[\s　]/g, '');
  const m = s.match(/^(\d{1,2})[\/月\-](\d{1,2})/);
  if (m) return { m: +m[1], d: +m[2] };
  return null;
}

function toDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const s = String(v).trim();
  const m = s.match(/^(\d{4})[-\/年](\d{1,2})[-\/月](\d{1,2})/);
  if (!m) return null;
  const d = new Date(+m[1], +m[2] - 1, +m[3]);
  return isNaN(d.getTime()) ? null : d;
}

// 「¥74,671」でも 74671 でも金額として読む
function readMoney_(v) {
  if (typeof v === 'number') return isNaN(v) ? 0 : Math.round(v);
  const s = String(v == null ? '' : v).replace(/[^0-9.\-]/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : Math.round(n);
}

function normYm_(v) {
  const m = String(v == null ? '' : v).match(/^(\d{4})-(\d{1,2})$/);
  if (!m) return '';
  return `${m[1]}-${('0' + m[2]).slice(-2)}`;
}

function thisYm_() {
  const d = new Date();
  return `${d.getFullYear()}-${('0' + (d.getMonth() + 1)).slice(-2)}`;
}

function prop_(name) {
  return String(PropertiesService.getScriptProperties().getProperty(name) || '').trim();
}

function cacheGet_(ym) {
  try {
    const s = CacheService.getScriptCache().get('remind_' + ym);
    return s ? JSON.parse(s) : null;
  } catch (e) { return null; }
}

function cachePut_(ym, data) {
  try { CacheService.getScriptCache().put('remind_' + ym, JSON.stringify(data), 21600); }
  catch (e) { /* 控えに置けなくても、その場で計算すれば出せる */ }
}
