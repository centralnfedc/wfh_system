/**
 * Google Apps Script backend for WFH System
 * Deploy as Web App (Execute as: Me, Who has access: Anyone)
 */

const SHEETS = {
  USERS: 'USERS',
  ATTENDANCE: 'ATTENDANCE',
  SETTINGS: 'SETTINGS',
  LOGS: 'LOGS',
};

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = (p.action || '').trim();

  try {
    let result;
    switch (action) {
      case 'getUsers':
        result = { status: 'ok', data: getUsers_() };
        break;
      case 'getTodayAttendance':
        result = { status: 'ok', data: getTodayAttendance_() };
        break;
      case 'getSettings':
        result = { status: 'ok', data: getSettings_() };
        break;
      case 'login':
        result = login_(p.username, p.password);
        break;
      default:
        result = { status: 'error', message: 'Unknown action' };
    }
    return output_(result, p.callback);
  } catch (err) {
    return output_({ status: 'error', message: String(err) }, p.callback);
  }
}

function doPost(e) {
  let payload = {};
  try {
    payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (_) {}

  try {
    const action = (payload.action || '').trim();
    switch (action) {
      case 'checkin':
      case 'checkout':
      case 'editRecord':
      case 'deleteRecord':
        saveAttendanceAction_(action, payload);
        break;
      case 'saveSettings':
        saveSettings_(payload);
        break;
      case 'deleteUser':
        deleteUser_(payload.username);
        break;
      case 'login':
        appendLog_('LOGIN', payload);
        break;
      default:
        appendLog_('UNKNOWN_POST', payload);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function output_(obj, callback) {
  const text = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + text + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function getUsers_() {
  const sh = getSheet_(SHEETS.USERS, ['id', 'username', 'password', 'firstName', 'lastName', 'position', 'dept', 'role', 'status']);
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  const head = rows[0];
  return rows.slice(1)
    .filter(r => String(r[1]).trim())
    .map(r => rowToObj_(head, r));
}

function getTodayAttendance_() {
  const sh = getSheet_(SHEETS.ATTENDANCE, ['id', 'date', 'username', 'location', 'checkIn', 'checkOut', 'task', 'gps', 'img', 'updatedAt']);
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];

  const tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const head = rows[0];

  return rows.slice(1)
    .map(r => rowToObj_(head, r))
    .filter(r => String(r.date) === today);
}

function getSettings_() {
  const sh = getSheet_(SHEETS.SETTINGS, ['key', 'value']);
  const rows = sh.getDataRange().getValues();
  const out = {};
  rows.slice(1).forEach(r => {
    const k = String(r[0] || '').trim();
    if (k) out[k] = r[1];
  });
  return out;
}

function login_(username, password) {
  const users = getUsers_();
  const found = users.find(u =>
    String(u.username).trim() === String(username || '').trim() &&
    String(u.password) === String(password || '') &&
    String(u.status || 'active').toLowerCase() !== 'inactive'
  );
  if (!found) return { status: 'error', message: 'invalid_credentials' };
  appendLog_('LOGIN', { username: found.username });
  return { status: 'ok', user: found };
}

function saveAttendanceAction_(action, payload) {
  const sh = getSheet_(SHEETS.ATTENDANCE, ['id', 'date', 'username', 'location', 'checkIn', 'checkOut', 'task', 'gps', 'img', 'updatedAt']);
  const rows = sh.getDataRange().getValues();
  const head = rows[0];
  const idx = indexMap_(head);

  if (action === 'deleteRecord') {
    const id = String(payload.recordId || '');
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][idx.id]) === id) sh.deleteRow(i + 1);
    }
    return;
  }

  const rowObj = {
    id: payload.id || payload.recordId || String(new Date().getTime()),
    date: payload.date || today_(),
    username: payload.username || '',
    location: payload.location || '',
    checkIn: payload.checkIn || '',
    checkOut: payload.checkOut || '',
    task: payload.task || '',
    gps: payload.gps || '',
    img: payload.img || '',
    updatedAt: new Date().toISOString(),
  };

  // upsert by id
  let rowNo = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idx.id]) === String(rowObj.id)) { rowNo = i + 1; break; }
  }

  const values = head.map(h => rowObj[h] !== undefined ? rowObj[h] : '');
  if (rowNo > 0) {
    sh.getRange(rowNo, 1, 1, head.length).setValues([values]);
  } else {
    sh.appendRow(values);
  }
}

function saveSettings_(payload) {
  const sh = getSheet_(SHEETS.SETTINGS, ['key', 'value']);
  const rows = sh.getDataRange().getValues();
  const map = {};
  rows.slice(1).forEach((r, i) => map[String(r[0])] = i + 2);

  Object.keys(payload).forEach(k => {
    if (['action', 'savedBy'].includes(k)) return;
    if (map[k]) {
      sh.getRange(map[k], 2).setValue(payload[k]);
    } else {
      sh.appendRow([k, payload[k]]);
    }
  });
  appendLog_('SAVE_SETTINGS', { savedBy: payload.savedBy || '' });
}

function deleteUser_(username) {
  const sh = getSheet_(SHEETS.USERS, ['id', 'username', 'password', 'firstName', 'lastName', 'position', 'dept', 'role', 'status']);
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][1]).trim() === String(username || '').trim()) sh.deleteRow(i + 1);
  }
}

function appendLog_(type, data) {
  const sh = getSheet_(SHEETS.LOGS, ['timestamp', 'type', 'data']);
  sh.appendRow([new Date().toISOString(), type, JSON.stringify(data || {})]);
}

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  return sh;
}

function rowToObj_(head, row) {
  const o = {};
  head.forEach((h, i) => o[h] = row[i]);
  return o;
}

function indexMap_(head) {
  const m = {};
  head.forEach((h, i) => m[h] = i);
  return m;
}

function today_() {
  const tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
}

function setupSampleData() {
  const users = getSheet_(SHEETS.USERS, ['id', 'username', 'password', 'firstName', 'lastName', 'position', 'dept', 'role', 'status']);
  if (users.getLastRow() === 1) {
    users.appendRow([1, 'admin', '1234', 'ผู้ดูแล', 'ระบบ', 'ผู้ดูแลระบบ', 'กลาง', 'admin', 'active']);
    users.appendRow([2, 'user1', '1234', 'สมชาย', 'ใจดี', 'เจ้าหน้าที่', 'งานบุคคล', 'user', 'active']);
  }
}
