/**
 * 飲み会調整 - Google Apps Script（スプレッドシートにバインドして使う）
 * シート「回答」: 1人1行。同じ名前は上書き。
 *   列：名前 / 更新日時 / ひとこと / 日付:希望月 / テーマ:◯◯ ...
 * シート「提案」: 参加者が追加した企画。列：企画 / 提案者 / 日時
 */
const SHEET_NAME = '回答';
const PROP_SHEET = '提案';
const FIXED = ['名前', '更新日時', 'ひとこと'];
const MAX_PROPOSALS = 12;

function doGet() {
  return json_({ ok: true, answers: readAll_(), proposals: readProps_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = JSON.parse(e.postData.contents);
    const name = String(p.name || '').trim().slice(0, 12);
    if (!name) return json_({ ok: false, error: '名前がありません' });

    // 企画の提案（あれば先に登録）
    addProps_(p.proposals || [], name);

    const sh = sheet_();
    const dateKeys = Object.keys(p.dates || {}).map(d => '日付:' + d);
    const themeList = (p.themes || []).map(String);

    // 見出しに無い列は右端に追加
    let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const need = dateKeys.concat(themeList.map(t => 'テーマ:' + t)).filter(h => headers.indexOf(h) < 0);
    if (need.length) {
      sh.getRange(1, headers.length + 1, 1, need.length).setValues([need]);
      headers = headers.concat(need);
    }

    const row = headers.map(h => {
      if (h === '名前') return name;
      if (h === '更新日時') return new Date();
      if (h === 'ひとこと') return String(p.comment || '').slice(0, 60);
      if (h.indexOf('日付:') === 0) return (p.dates || {})[h.slice(3)] || '';
      if (h.indexOf('テーマ:') === 0) return themeList.indexOf(h.slice(4)) >= 0 ? '♥' : '';
      return '';
    });

    // 同名があれば上書き、なければ追加
    const last = sh.getLastRow();
    const names = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
    const idx = names.indexOf(name);
    if (idx >= 0) sh.getRange(idx + 2, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);

    return json_({ ok: true, answers: readAll_(), proposals: readProps_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function readAll_() {
  const sh = sheet_();
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  return values.filter(r => r[0] !== '').map(r => {
    const a = { name: String(r[0]), updatedAt: new Date(r[1]).toISOString(), comment: String(r[2] || ''), dates: {}, themes: [] };
    headers.forEach((h, i) => {
      if (h.indexOf('日付:') === 0 && r[i]) a.dates[h.slice(3)] = String(r[i]);
      if (h.indexOf('テーマ:') === 0 && r[i]) a.themes.push(h.slice(4));
    });
    return a;
  });
}

function readProps_() {
  const sh = propSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 2).getValues()
    .filter(r => r[0] !== '')
    .map(r => ({ name: String(r[0]), by: String(r[1]) }));
}

function addProps_(list, by) {
  const sh = propSheet_();
  const existing = readProps_().map(x => x.name);
  list.map(s => String(s).trim().slice(0, 15)).filter(Boolean).forEach(t => {
    if (existing.indexOf(t) >= 0 || existing.length >= MAX_PROPOSALS) return;
    sh.appendRow([t, by, new Date()]);
    existing.push(t);
  });
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, FIXED.length).setValues([FIXED]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function propSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PROP_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PROP_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['企画', '提案者', '日時']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
