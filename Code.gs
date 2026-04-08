// שם הגיליון שבו נשמרות עסקאות היתרות (Sheet tab name)
const SHEET_NAME = 'Data';
// גיליון שחקנים: עמודה A בלבד — שם (כותרת: שחקן)
const PLAYERS_SHEET_NAME = 'Players';
// גיליון Debt: יומן טרנזקציות בלבד — שורה לכל עדכון (תאריך, שחקן, דלתא כמות, דלתא סכום)
const DEBT_SHEET_NAME = 'Debt';

const DEBT_HEADERS_NEW_ = ['תאריך ושעה', 'שחקן', 'דלתא כמות', 'דלתא סכום'];

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
 * ?action=getPlayers — רשימת שחקנים מ-Players + סיכום מ-Debt (אופציונלי: fromYm, toYm בפורמט YYYY-MM)
 * אחרת — מחזיר totals נוכחי מגיליון Data
 */
function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action ? e.parameter.action : '';

    if (action === 'getPlayers') {
      return getPlayers_(e && e.parameter ? e.parameter : {});
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

function getOrCreatePlayersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PLAYERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PLAYERS_SHEET_NAME);
    sheet.getRange(1, 1).setValue('שחקן');
  }
  return sheet;
}

function getOrCreateDebtSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DEBT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DEBT_SHEET_NAME);
    sheet.getRange(1, 1, 1, 4).setValues([DEBT_HEADERS_NEW_]);
  }
  return sheet;
}

/**
 * פורמט ישן: שורה 1 — שחקן | כמות | סכום | עדכון אחרון
 */
function isLegacyDebtSheet_(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return false;
  const h = sheet.getRange(1, 1, 1, 4).getValues()[0];
  const a = String(h[0] || '').trim();
  const b = String(h[1] || '').trim();
  return a === 'שחקן' && b === 'כמות';
}

function legacyDebtErrorResponse_() {
  return jsonResponse_({
    success: false,
    error: 'גיליון Debt בפורמט ישן (שורה לשחקן עם כמות/סכום מצטברים). יש לגבות, ליצור גיליון Debt חדש עם הכותרות: תאריך ושעה, שחקן, דלתא כמות, דלתא סכום, להעביר שמות לגיליון Players, ולפרוס מחדש.'
  });
}

function cellToDate_(value) {
  if (value === '' || value === null || value === undefined) {
    return null;
  }
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }
  if (typeof value === 'string' && value.trim() !== '') {
    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }
  }
  return null;
}

/**
 * ym: "YYYY-MM" → { y, m } (m = 1..12)
 */
function parseYm_(ym) {
  if (!ym || typeof ym !== 'string') return null;
  const parts = ym.trim().split('-');
  if (parts.length !== 2) return null;
  const y = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10);
  if (isNaN(y) || isNaN(m) || m < 1 || m > 12) return null;
  return { y: y, m: m };
}

function monthStartDate_(y, m) {
  return new Date(y, m - 1, 1, 0, 0, 0, 0);
}

function monthEndDate_(y, m) {
  return new Date(y, m, 0, 23, 59, 59, 999);
}

/**
 * מחזיר { start, end } כ-Dates כוללים, או null אם לא תקין
 */
function parseYmRange_(fromYm, toYm) {
  const from = parseYm_(fromYm);
  const to = parseYm_(toYm);
  if (!from || !to) return null;
  const start = monthStartDate_(from.y, from.m);
  const end = monthEndDate_(to.y, to.m);
  if (start.getTime() > end.getTime()) return null;
  return { start: start, end: end };
}

function playerExistsInPlayersSheet_(targetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PLAYERS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return false;
  const lastRow = sheet.getLastRow();
  const col = sheet.getRange(2, 1, lastRow, 1).getValues();
  const target = String(targetName).trim();
  for (let i = 0; i < col.length; i++) {
    const name = col[i][0];
    if (name === '' || name === null || name === undefined) continue;
    if (String(name).trim() === target) return true;
  }
  return false;
}

function getPlayerNamesOrdered_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PLAYERS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const lastRow = sheet.getLastRow();
  const col = sheet.getRange(2, 1, lastRow, 1).getValues();
  const names = [];
  for (let i = 0; i < col.length; i++) {
    const name = col[i][0];
    if (name === '' || name === null || name === undefined) continue;
    names.push(String(name).trim());
  }
  return names;
}

/**
 * רשומות { when, c, a } — רק שורות מזווגות (כמות+סכום).
 * N = מספר העדכונים; k = N % 5, ואם 0 אז k=5. מסכמים את k השורות האחרונות לפי תאריך.
 */
function aggregatePairedDebtMod5_(pairedEntries) {
  if (!pairedEntries || pairedEntries.length === 0) {
    return { count: 0, amount: 0, lastUpdated: null };
  }
  const sorted = pairedEntries.slice().sort(function (x, y) {
    return x.when.getTime() - y.when.getTime();
  });
  const N = sorted.length;
  var k = N % 5;
  if (k === 0) {
    k = 5;
  }
  var start = N - k;
  var count = 0;
  var amount = 0;
  var lastUpdated = null;
  for (var i = start; i < N; i++) {
    count += sorted[i].c;
    amount += sorted[i].a;
    var w = sorted[i].when;
    if (w && !isNaN(w.getTime())) {
      if (!lastUpdated || w.getTime() > lastUpdated.getTime()) {
        lastUpdated = w;
      }
    }
  }
  return { count: count, amount: amount, lastUpdated: lastUpdated };
}

/**
 * params: אובייקט מ-e.parameter — action, fromYm, toYm (אופציונלי)
 */
function getPlayers_(params) {
  const p = params || {};
  const fromYm = p.fromYm ? String(p.fromYm).trim() : '';
  const toYm = p.toYm ? String(p.toYm).trim() : '';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const debtSheet = ss.getSheetByName(DEBT_SHEET_NAME);

  if (debtSheet && isLegacyDebtSheet_(debtSheet)) {
    return legacyDebtErrorResponse_();
  }

  let rangeFilter = null;
  if (fromYm !== '' || toYm !== '') {
    if (fromYm === '' || toYm === '') {
      return jsonResponse_({
        success: false,
        error: 'יש לשלוח גם fromYm וגם toYm (פורמט YYYY-MM), או להשאיר את שניהם ריקים לכל התקופה'
      });
    }
    rangeFilter = parseYmRange_(fromYm, toYm);
    if (!rangeFilter) {
      return jsonResponse_({
        success: false,
        error: 'טווח תאריכים לא תקין: ודא פורמט YYYY-MM ושהחודש "מ-" אינו אחרי "עד"'
      });
    }
  }

  const playerNames = getPlayerNamesOrdered_();
  const pairedByPlayer = {};

  if (debtSheet && debtSheet.getLastRow() >= 2) {
    const lastRow = debtSheet.getLastRow();
    const data = debtSheet.getRange(2, 1, lastRow, 4).getValues();
    for (let r = 0; r < data.length; r++) {
      const row = data[r];
      const when = cellToDate_(row[0]);
      const playerName = row[1];
      if (playerName === '' || playerName === null || playerName === undefined) continue;
      const name = String(playerName).trim();
      const c = Number(row[2]) || 0;
      const a = Number(row[3]) || 0;

      if (c === 0 || a === 0) {
        continue;
      }

      if (!when || isNaN(when.getTime())) {
        continue;
      }

      if (rangeFilter) {
        const t = when.getTime();
        if (t < rangeFilter.start.getTime() || t > rangeFilter.end.getTime()) continue;
      }

      if (!pairedByPlayer[name]) {
        pairedByPlayer[name] = [];
      }
      pairedByPlayer[name].push({ when: when, c: c, a: a });
    }
  }

  const players = [];
  for (let i = 0; i < playerNames.length; i++) {
    const nm = playerNames[i];
    const agg = aggregatePairedDebtMod5_(pairedByPlayer[nm]);
    const row = {
      name: nm,
      count: agg.count,
      amount: agg.amount
    };
    if (agg.lastUpdated) {
      row.lastUpdated = agg.lastUpdated.toISOString();
    }
    players.push(row);
  }

  return jsonResponse_({ success: true, players: players });
}

function addPlayer_(name) {
  if (!name || String(name).trim() === '') {
    return jsonResponse_({ success: false, error: 'Missing player name' });
  }

  const sheet = getOrCreatePlayersSheet_();
  const trimmed = String(name).trim();
  const lastRow = sheet.getLastRow();
  const effectiveLast = lastRow < 2 ? 1 : lastRow;

  for (let i = 2; i <= effectiveLast; i++) {
    if (String(sheet.getRange(i, 1).getValue()).trim() === trimmed) {
      return jsonResponse_({ success: false, error: 'Player already exists' });
    }
  }

  sheet.appendRow([trimmed]);

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
  if (c !== 0 && a === 0) {
    return jsonResponse_({
      success: false,
      error: 'לא ניתן לעדכן כמות בלי סכום — יש לצרף דלתא סכום (למשל מהשדה לפני לחיצה על +1)'
    });
  }

  const target = String(name).trim();
  if (!playerExistsInPlayersSheet_(target)) {
    return jsonResponse_({ success: false, error: 'Player not found' });
  }

  const debtSheet = getOrCreateDebtSheet_();
  if (isLegacyDebtSheet_(debtSheet)) {
    return legacyDebtErrorResponse_();
  }

  const now = new Date();
  debtSheet.appendRow([now, target, c, a]);

  return jsonResponse_({ success: true });
}

function deletePlayer_(name) {
  if (!name || String(name).trim() === '') {
    return jsonResponse_({ success: false, error: 'Missing player name' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PLAYERS_SHEET_NAME);
  if (!sheet) {
    return jsonResponse_({ success: false, error: 'Players sheet not found' });
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse_({ success: false, error: 'Player not found' });
  }

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
