/**
 * Team Bill Splitter – Google Sheet backend
 *
 * 1. Run setup() once (it creates the sheets and adds your team).
 * 2. Anyone with the app link can see and add bills.
 *    Only someone with ADMIN_PIN can change names or UPI IDs, or restore a backup.
 * 3. Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
 * 4. Paste the Web app URL into SCRIPT_URL at the top of index.html.
 */

const ADMIN_PIN = '2001';

// Your team, added to the People sheet the first time you run setup().
// Check the spellings and UPI IDs before running it.
const STARTING_TEAM = [
  ['Aditya Sunit Kanoi', 'adityakanoi123@oksbi'],
  ['Kiran Senthilkumar', 'kiran.in2002-1@okicici'],
  ['Arshavi', 'arshavi110701-1@okaxis'],
  ['Sandeep Singh Bijral', 'sandeepbijral123-4@okaxis'],
  ['Chaithra Gowda', 'chaithra.papu@okicici'],
  ['Raman Rai', 'ramanbnglr@okicici'],
  ['Bhumika Jain', ''],
  ['Pranav Garg', 'pranavgarg5570@oksbi'],
  ['Sachidananda Tripathy', '6303312125@ybl']
];

const P_HEAD = ['id', 'name', 'upi'];
const B_HEAD = ['id', 'date', 'kind', 'title', 'amount', 'paid_by_id', 'paid_by', 'member_ids', 'members', 'birthday_id', 'to_id', 'added_at'];
const KINDS = ['meal', 'birthday', 'other', 'settle'];

function setup() {
  const people = sheet_('People', P_HEAD);
  sheet_('Bills', B_HEAD);
  if (people.getLastRow() <= 1 && STARTING_TEAM.length) {
    const rows = STARTING_TEAM.map(function (t, i) { return ['p' + Date.now() + '-' + i, t[0], t[1]]; });
    people.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  Logger.log('Setup done.');
}

function doGet(e) {
  return json_(Object.assign({ ok: true }, readAll_()));
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad request' }); }
  const adminOps = ['upsertPerson', 'removePerson', 'import', 'checkAdmin'];
  if (adminOps.indexOf(body.op) >= 0 && String(body.adminPin || '') !== ADMIN_PIN) return json_({ ok: false, error: 'admin' });
  if (body.op === 'checkAdmin') return json_(Object.assign({ ok: true }, readAll_()));
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    applyOp_(body);
    SpreadsheetApp.flush();
    return json_(Object.assign({ ok: true }, readAll_()));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function applyOp_(b) {
  if (b.op === 'addBill') return addBill_(b.bill);
  if (b.op === 'deleteBill') return deleteRow_(sheet_('Bills', B_HEAD), String(b.id));
  if (b.op === 'upsertPerson') return upsertPerson_(b.person);
  if (b.op === 'removePerson') return deleteRow_(sheet_('People', P_HEAD), String(b.id));
  if (b.op === 'import') return importAll_(b.people || [], b.bills || []);
  throw new Error('Unknown action');
}

// ---------- people ----------
function upsertPerson_(p) {
  if (!p || !p.id || !p.name) throw new Error('Person needs a name');
  const sh = sheet_('People', P_HEAD);
  const row = [txt_(p.id), txt_(p.name), txt_(p.upi || '')];
  const r = findRow_(sh, row[0]);
  if (r) sh.getRange(r, 1, 1, 3).setValues([row]);
  else sh.appendRow(row);
}

// ---------- bills ----------
function billRow_(b, names) {
  if (!b || !b.id) throw new Error('Bill needs an id');
  const kind = KINDS.indexOf(b.kind) >= 0 ? b.kind : 'other';
  const members = (b.members || []).map(String);
  const nm = function (id) { return names[id] || ''; };
  return [
    txt_(b.id), txt_(b.date || ''), kind, txt_(b.title || ''), Number(b.amt) || 0,
    txt_(b.payer || ''), txt_(nm(b.payer)), members.join(','), txt_(members.map(nm).join(', ')),
    txt_(b.bday || ''), txt_(b.to || ''), new Date()
  ];
}

function addBill_(b) {
  const sh = sheet_('Bills', B_HEAD);
  if (findRow_(sh, String(b && b.id))) return; // already saved
  sh.appendRow(billRow_(b, names_()));
}

function importAll_(people, bills) {
  const ps = sheet_('People', P_HEAD), bs = sheet_('Bills', B_HEAD);
  if (ps.getLastRow() > 1) ps.getRange(2, 1, ps.getLastRow() - 1, P_HEAD.length).clearContent();
  if (bs.getLastRow() > 1) bs.getRange(2, 1, bs.getLastRow() - 1, B_HEAD.length).clearContent();
  const prow = people.filter(function (p) { return p && p.id && p.name; })
    .map(function (p) { return [txt_(p.id), txt_(p.name), txt_(p.upi || '')]; });
  if (prow.length) ps.getRange(2, 1, prow.length, 3).setValues(prow);
  const names = {};
  prow.forEach(function (r) { names[r[0]] = r[1]; });
  const brow = bills.filter(function (b) { return b && b.id; }).map(function (b) { return billRow_(b, names); });
  if (brow.length) bs.getRange(2, 1, brow.length, B_HEAD.length).setValues(brow);
}

// ---------- reading ----------
function readAll_() {
  const people = rows_(sheet_('People', P_HEAD)).filter(function (r) { return r[0] !== ''; }).map(function (r) {
    return { id: String(r[0]), name: String(r[1]), upi: String(r[2] || '') };
  });
  const bills = rows_(sheet_('Bills', B_HEAD)).filter(function (r) { return r[0] !== ''; }).map(function (r) {
    const b = {
      id: String(r[0]), date: date_(r[1]), kind: String(r[2] || 'other'), title: String(r[3] || ''),
      amt: Number(r[4]) || 0, payer: String(r[5] || ''),
      members: String(r[7] || '').split(',').filter(function (x) { return x; })
    };
    if (r[9]) b.bday = String(r[9]);
    if (r[10]) b.to = String(r[10]);
    return b;
  });
  return { people: people, bills: bills };
}

// ---------- helpers ----------
function sheet_(name, head) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
    if (name === 'Bills') sh.getRange('B:B').setNumberFormat('@');
    if (name === 'People') sh.getRange('C:C').setNumberFormat('@');
  }
  return sh;
}
function rows_(sh) { const v = sh.getDataRange().getValues(); v.shift(); return v; }
function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return 0;
}
function deleteRow_(sh, id) { const r = findRow_(sh, id); if (r) sh.deleteRow(r); }
function names_() {
  const n = {};
  rows_(sheet_('People', P_HEAD)).forEach(function (r) { n[String(r[0])] = String(r[1]); });
  return n;
}
function date_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}
// Keeps text as text, and stops anything typed from being read as a spreadsheet formula.
function txt_(v) {
  const s = String(v == null ? '' : v).slice(0, 300);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
