// Student Planner: Google Apps Script backend.
// Your planner items are saved in a Google Sheet called "Student Planner data"
// in your Google Drive. It is created automatically the first time you open the app.

const HEADERS = ['id', 'kind', 'title', 'date', 'time', 'done', 'created'];
const KINDS = ['assignment', 'turnin', 'test', 'extra', 'todo'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Student Planner')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function sheet_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('SHEET_ID');
  let ss = null;
  if (savedId) {
    try { ss = SpreadsheetApp.openById(savedId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('Student Planner data');
    props.setProperty('SHEET_ID', ss.getId());
    const sh = ss.getSheets()[0];
    sh.setName('Items');
    sh.getRange('A:G').setNumberFormat('@'); // keep dates and times as plain text
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return ss.getSheetByName('Items') || ss.getSheets()[0];
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function clean_(item) {
  const out = {};
  if ('id' in item) out.id = String(item.id).slice(0, 64);
  if ('kind' in item) out.kind = KINDS.indexOf(item.kind) >= 0 ? item.kind : 'todo';
  if ('title' in item) out.title = String(item.title).slice(0, 500);
  if ('date' in item) out.date = /^\d{4}-\d{2}-\d{2}$/.test(item.date) ? item.date : '';
  if ('time' in item) out.time = /^\d{2}:\d{2}$/.test(item.time) ? item.time : '';
  if ('done' in item) out.done = item.done === true;
  if ('created' in item) out.created = Number(item.created) || Date.now();
  return out;
}

function toRow_(item) {
  return HEADERS.map(function (h) { return item[h] === undefined ? '' : String(item[h]); });
}

function fromRow_(row) {
  return {
    id: String(row[0]), kind: String(row[1]), title: String(row[2]),
    date: String(row[3]), time: String(row[4]),
    done: row[5] === true || String(row[5]).toUpperCase() === 'TRUE',
    created: Number(row[6]) || 0,
  };
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

function getItems() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(fromRow_);
}

function addItem(item) {
  const it = clean_(item);
  if (!it.id || !it.title) throw new Error('An item needs a title.');
  return withLock_(function () {
    const sh = sheet_();
    if (findRow_(sh, it.id) < 0) sh.appendRow(toRow_(it));
    return true;
  });
}

function updateItem(id, patch) {
  return withLock_(function () {
    const sh = sheet_();
    const r = findRow_(sh, String(id));
    if (r < 0) return false;
    const cur = fromRow_(sh.getRange(r, 1, 1, HEADERS.length).getValues()[0]);
    const p = clean_(patch);
    delete p.id;
    sh.getRange(r, 1, 1, HEADERS.length).setValues([toRow_(Object.assign(cur, p))]);
    return true;
  });
}

function deleteItem(id) {
  return withLock_(function () {
    const sh = sheet_();
    const r = findRow_(sh, String(id));
    if (r > 0) sh.deleteRow(r);
    return true;
  });
}
