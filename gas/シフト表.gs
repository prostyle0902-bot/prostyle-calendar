// ===== 予測表に「シフト表」タブを自動で作る =====
// 現場管理カレンダーで入れた「現場・担当スタッフ・休み」から、
// 各月の予測表（例：2026年10月_予測表）に「シフト表」タブを作って書き込む。
// 縦：日付　横：社員　マス：その日に入る現場（または休み）
//
// 使い方（最初に1回だけ）
//   1. このファイルを「現場管理カレンダー予測表転記システム」に追加して保存
//   2. 上の関数一覧で setupShiftTrigger を選んで「実行」→ 権限を許可
//   → 以後10分ごとに自動で更新される。すぐ更新したいときは updateShiftTables を実行
//
// ※ シフト表タブは毎回まるごと書き直すので、手で書き込んだ内容は消えます

const SHIFT_TAB_NAME = 'シフト表';

// シフト表に並べる社員（カレンダーアプリ index.html の STAFF_OWN と同じ順番）
const SHIFT_STAFF = ['平野健雄','安藤正樹','伏田幸枝','小林俊樹','石毛貴英','池田久史','三浦達弥','飯田怜飛','飯田栄'];

// 休みの種類ごとの書き方と色
const SHIFT_OFF = {
  holiday:  { label: '休',     color: '#fee2e2' },
  half:     { label: '半休',   color: '#fff1f1' },
  special:  { label: '特休',   color: '#e2e8f0' },
  paid:     { label: '有給',   color: '#d4edda' },
  halfPaid: { label: '有給半', color: '#eef8f1' },
  kibou:    { label: '希望休', color: '#ffe9c7' },
  refresh:  { label: 'リフレ', color: '#ede7f6' },
};

// 10分ごとに updateShiftTables を動かす設定（最初に1回だけ実行）
function setupShiftTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'updateShiftTables')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('updateShiftTables').timeBased().everyMinutes(10).create();
  updateShiftTables();
}

// 先月〜再来月の予測表のシフト表を作り直す（予測表がない月は飛ばす）
function updateShiftTables() {
  const data = {
    events:     getAllEvents(),
    holidays:   normalizeDateKeys_(getAllHolidays()),
    paidLeaves: normalizeDateKeys_(getAllPaidLeaves()),
    kibou:      normalizeDateKeys_(getAllKibou()),
    refresh:    normalizeDateKeys_(getAllRefresh()),
  };
  data.events = normalizeDateKeys_(data.events);

  const now = new Date();
  for (let diff = -1; diff <= 2; diff++) {
    const d = new Date(now.getFullYear(), now.getMonth() + diff, 1);
    const y = d.getFullYear(), m = d.getMonth() + 1;
    const files = DriveApp.getFilesByName(`${y}年${m}月_予測表`);
    if (!files.hasNext()) continue;
    try {
      const ss = SpreadsheetApp.openById(files.next().getId());
      writeShiftTable_(ss, y, m, data);
    } catch (e) {
      Logger.log(`シフト表の書き込みエラー（${y}年${m}月）: ` + e.toString());
    }
  }
}

// 「2026-10-5」のような日付も「2026-10-05」にそろえる
function normalizeDateKey_(s) {
  const p = String(s).split('-');
  if (p.length < 3) return String(s);
  return `${p[0]}-${('0' + parseInt(p[1], 10)).slice(-2)}-${('0' + parseInt(p[2], 10)).slice(-2)}`;
}

function normalizeDateKeys_(byDate) {
  const out = {};
  Object.keys(byDate || {}).forEach(k => {
    const key = normalizeDateKey_(k);
    out[key] = (out[key] || []).concat(byDate[k]);
  });
  return out;
}

// 休みの保存場所には「名前(半)」「名前(特:忌引き)」の形でも入っている
function parseOffName_(v) {
  const sp = String(v).match(/^(.+)\(特:([^)]+)\)$/);
  if (sp) return { name: sp[1], special: sp[2] };
  if (String(v).slice(-3) === '(半)') return { name: String(v).slice(0, -3), half: true };
  return { name: String(v) };
}

// その日のその人の休み（なければ null）
function offFor_(data, key, name) {
  for (const v of data.holidays[key] || []) {
    const o = parseOffName_(v);
    if (o.name !== name) continue;
    if (o.special) return { label: o.special === 'その他' ? SHIFT_OFF.special.label : o.special, color: SHIFT_OFF.special.color, full: true };
    if (o.half) return { label: SHIFT_OFF.half.label, color: SHIFT_OFF.half.color, full: false };
    return { label: SHIFT_OFF.holiday.label, color: SHIFT_OFF.holiday.color, full: true };
  }
  for (const v of data.paidLeaves[key] || []) {
    const o = parseOffName_(v);
    if (o.name !== name) continue;
    return o.half ? { label: SHIFT_OFF.halfPaid.label, color: SHIFT_OFF.halfPaid.color, full: false }
                  : { label: SHIFT_OFF.paid.label, color: SHIFT_OFF.paid.color, full: true };
  }
  if ((data.kibou[key] || []).indexOf(name) >= 0) return { label: SHIFT_OFF.kibou.label, color: SHIFT_OFF.kibou.color, full: true };
  if ((data.refresh[key] || []).indexOf(name) >= 0) return { label: SHIFT_OFF.refresh.label, color: SHIFT_OFF.refresh.color, full: true };
  return null;
}

// 1マスの中身と色。現場は「出社時間・作業時間・現場名」を1行ずつ
function shiftCell_(data, key, name) {
  const lines = [];
  const off = offFor_(data, key, name);
  if (off) lines.push(off.label);

  const evs = (data.events[key] || [])
    .filter(ev => (Array.isArray(ev.staff) ? ev.staff : []).indexOf(name) >= 0)
    .sort((a, b) => String(a.departure || a.start || '99').localeCompare(String(b.departure || b.start || '99')));
  evs.forEach(ev => {
    const time = ev.start ? (ev.end ? `${ev.start}〜${ev.end}` : `${ev.start}〜`) : '';
    const dep = ev.departure ? `${ev.departure}出社 ` : '';
    lines.push(`${dep}${time} ${ev.place}`.trim());
  });

  let color = null;
  if (off && (off.full || !evs.length)) color = off.color;
  return { text: lines.join('\n'), color: color };
}

// 1か月分の表（見出しと中身）を作る
function buildShiftGrid_(y, m, data) {
  const days = ['日','月','火','水','木','金','土'];
  const lastDay = new Date(y, m, 0).getDate();
  const values = [], colors = [];
  for (let d = 1; d <= lastDay; d++) {
    const key = `${y}-${('0' + m).slice(-2)}-${('0' + d).slice(-2)}`;
    const dow = new Date(y, m - 1, d).getDay();
    const rowColor = dow === 0 ? '#fdecec' : dow === 6 ? '#eaf3fc' : null;
    const row = [`${m}/${d}`, days[dow]], rowColors = [rowColor, rowColor];
    SHIFT_STAFF.forEach(name => {
      const c = shiftCell_(data, key, name);
      row.push(c.text);
      rowColors.push(c.color || rowColor);
    });
    values.push(row);
    colors.push(rowColors);
  }
  return { values: values, colors: colors };
}

function writeShiftTable_(ss, y, m, data) {
  let sheet = ss.getSheetByName(SHIFT_TAB_NAME);
  if (!sheet) sheet = ss.insertSheet(SHIFT_TAB_NAME);

  const grid = buildShiftGrid_(y, m, data);
  const cols = SHIFT_STAFF.length + 2;
  const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm');

  sheet.clear();
  sheet.getRange(1, 1).setValue(`${y}年${m}月 業務課シフト表（現場管理カレンダーから自動作成・${stamp} 更新・手で書き込まないでください）`);
  sheet.getRange(2, 1, 1, cols).setValues([['日付', '曜日'].concat(SHIFT_STAFF)])
    .setFontWeight('bold').setBackground('#1a1a1a').setFontColor('#ffffff').setHorizontalAlignment('center');

  const body = sheet.getRange(3, 1, grid.values.length, cols);
  body.setNumberFormat('@');
  body.setValues(grid.values);
  body.setBackgrounds(grid.colors);
  body.setWrap(true).setVerticalAlignment('top').setFontSize(10);
  body.setBorder(true, true, true, true, true, true, '#d0d0d0', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(3, 1, grid.values.length, 2).setHorizontalAlignment('center');

  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(2);
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 36);
  sheet.setColumnWidths(3, SHIFT_STAFF.length, 150);
}
