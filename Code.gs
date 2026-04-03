// שם הגיליון שבו נשמרות עסקאות היתרות (Sheet tab name)
const SHEET_NAME = 'Data';
// שם הגיליון לניהול שחקנים: עמודה A שם, B כמות, C סכום
const DEBT_SHEET_NAME = 'Debt';

/**
 * קריאה ב-POST מה-HTML:
 * ללא action (או action ריק): { "user": "מאור" | "עידו", "amount": מספר, "note": "..." }
 * עם action: addPlayer | updatePlayer | deletePlayer
 */
function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) {
      return jsonResponse_({
        success: false,
        error: 'Missing post data'
      });
    }

    const data = JSON.parse(e.postData.contents);
    const action = data.action || '';

    if (action === 'addPlayer') {
      return addPlayer_(data.name);
    }
    if (action === 'updatePlayer') {
      return updatePlayer_(data.name, data.count, data.amount);
    }
    if (action === 'deletePlayer') {
      return deletePlayer_(data.name);
    }

    const user = data.user;
    const amount = Number(data.amount);
    const note = data.note || '';

    if (!user || isNaN(amount)) {
      return jsonResponse_({
        success: false,
        error: 'Invalid user or amount'
      });
    }

    const allowedUsers = ['מאור', 'עידו'];
    if (allowedUsers.indexOf(user) === -1) {
      return jsonResponse_({
        success: false,
        error: 'Invalid user'
      });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return jsonResponse_({
        success: false,
        error: 'Sheet "' + SHEET_NAME + '" not found'
      });
    }

    sheet.appendRow([new Date(), user, amount, note]);

    const totals = calculateTotals_(sheet);

    return jsonResponse_({
      success: true,
      totals: totals
    });
  } catch (err) {
    const msg = err && err.message ? err.message : err;
    return jsonResponse_({
      success: false,
      error: msg
    });
  }
}

/**
 * קריאה ב-GET מה-HTML:
 * ?action=getPlayers — מחזיר רשימת שחקנים מגיליון Debt
 * אחרת — מחזיר totals נוכחי מגיליון Data
 */
function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action ? e.parameter.action : '';

    if (action === 'getPlayers') {
      return getPlayers_();
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return jsonResponse_({
        success: false,
        error: 'Sheet "' + SHEET_NAME + '" not found'
      });
    }

    const totals = calculateTotals_(sheet);

    return jsonResponse_({
      success: true,
      totals: totals
    });
  } catch (err) {
    const msg = err && err.message ? err.message : err;
    return jsonResponse_({
      success: false,
      error: msg
    });
  }
}

/**
 * סיכומים למאור ולעידו מתוך גיליון Data
 * עמודה A: Timestamp
 * עמודה B: User
 * עמודה C: Amount
 * עמודה D: Note
 */
function calculateTotals_(sheet) {
  const values = sheet.getDataRange().getValues();

  const totals = {
    'מאור': 0,
    'עידו': 0
  };

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const user = row[1];
    const amount = Number(row[2]) || 0;

    if (totals.hasOwnProperty(user)) {
      totals[user] += amount;
    }
  }

  return totals;
}

function getOrCreateDebtSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DEBT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DEBT_SHEET_NAME);
    sheet.getRange(1, 1, 1, 3).setValues([['שחקן', 'כמות', 'סכום']]);
  }
  return sheet;
}

function getPlayers_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DEBT_SHEET_NAME);

  if (!sheet) {
    return jsonResponse_({ success: true, players: [] });
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse_({ success: true, players: [] });
  }

  const data = sheet.getRange(2, 1, lastRow, 3).getValues();
  const players = [];
  for (let i = 0; i < data.length; i++) {
    const name = data[i][0];
    if (name === '' || name === null || name === undefined) continue;
    players.push({
      name: String(name),
      count: Number(data[i][1]) || 0,
      amount: Number(data[i][2]) || 0
    });
  }

  return jsonResponse_({ success: true, players: players });
}

function addPlayer_(name) {
  if (!name || String(name).trim() === '') {
    return jsonResponse_({ success: false, error: 'Missing player name' });
  }

  const sheet = getOrCreateDebtSheet_();
  const trimmed = String(name).trim();
  const lastRow = sheet.getLastRow();

  for (let i = 2; i <= lastRow; i++) {
    if (String(sheet.getRange(i, 1).getValue()).trim() === trimmed) {
      return jsonResponse_({ success: false, error: 'Player already exists' });
    }
  }

  sheet.appendRow([trimmed, 0, 0]);

  return jsonResponse_({ success: true });
}

function updatePlayer_(name, countDelta, amountDelta) {
  if (!name || String(name).trim() === '') {
    return jsonResponse_({ success: false, error: 'Missing player name' });
  }

  const c = Number(countDelta) || 0;
  const a = Number(amountDelta) || 0;
  if (c === 0 && a === 0) {
    return jsonResponse_({ success: false, error: 'Nothing to add' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DEBT_SHEET_NAME);
  if (!sheet) {
    return jsonResponse_({ success: false, error: 'Debt sheet not found' });
  }

  const lastRow = sheet.getLastRow();
  const target = String(name).trim();

  for (let i = 2; i <= lastRow; i++) {
    if (String(sheet.getRange(i, 1).getValue()).trim() === target) {
      const currentCount = Number(sheet.getRange(i, 2).getValue()) || 0;
      const currentAmount = Number(sheet.getRange(i, 3).getValue()) || 0;
      sheet.getRange(i, 2).setValue(currentCount + c);
      sheet.getRange(i, 3).setValue(currentAmount + a);
      return jsonResponse_({ success: true });
    }
  }

  return jsonResponse_({ success: false, error: 'Player not found' });
}

function deletePlayer_(name) {
  if (!name || String(name).trim() === '') {
    return jsonResponse_({ success: false, error: 'Missing player name' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DEBT_SHEET_NAME);
  if (!sheet) {
    return jsonResponse_({ success: false, error: 'Debt sheet not found' });
  }

  const lastRow = sheet.getLastRow();
  const target = String(name).trim();

  for (let i = 2; i <= lastRow; i++) {
    if (String(sheet.getRange(i, 1).getValue()).trim() === target) {
      sheet.deleteRow(i);
      return jsonResponse_({ success: true });
    }
  }

  return jsonResponse_({ success: false, error: 'Player not found' });
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
