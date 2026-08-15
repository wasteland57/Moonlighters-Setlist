/**
 * Setlist Builder — Google Sheets backend
 *
 * Bind this to the same spreadsheet that holds your song library:
 *   Extensions -> Apps Script, paste this in, Save.
 *   Deploy -> New deployment -> type: Web app
 *     Execute as:      Me
 *     Who has access:  Anyone
 *   Copy the /exec URL into the Setlist Builder's "Web app URL" box.
 *
 * "Anyone" means anyone holding that URL can read and write the Setlists tab.
 * Set TOKEN below to a random string and put the same string in the app if you
 * want a basic guard against someone who stumbles onto the URL.
 *
 * Writes go to a tab named Setlists, one row per song, so the data stays
 * readable and filterable in the sheet itself. Your song library tab is never
 * touched.
 */

var SHEET_NAME = 'Setlists';
var TOKEN = 'rFRPN1FIBcTV3loeWt5c_kYP';   // must match the ?token= in moonlighters.html
var HEADERS = ['ID','Name','Venue','Date','Set','Position','Title','Length','Notes','Updated'];

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    if (TOKEN && p.token !== TOKEN) throw new Error('Not authorised');
    switch (p.action) {
      case 'list':   out = { ok: true, items: listSetlists() }; break;
      case 'load':   out = { ok: true, setlist: loadSetlist(p.id) }; break;
      case 'save':
        var saved = saveSetlist(JSON.parse(p.data));
        out = { ok: true, id: saved.id, rows: saved.rows };
        break;
      case 'delete': out = { ok: true, removed: deleteSetlist(p.id) }; break;
      case 'ping':   out = { ok: true, sheet: sheet().getName(), rows: Math.max(0, sheet().getLastRow() - 1) }; break;
      default: throw new Error('Unknown action: ' + p.action);
    }
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  var json = JSON.stringify(out);
  if (p.callback) {
    return ContentService
      .createTextOutput(p.callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  ensureHeaders(sh);
  return sh;
}

/**
 * Every read starts at row 2, so row 1 MUST be the header. If the tab already
 * existed when this script first ran, no header was written and the first song
 * landed in row 1 — invisible to every read and impossible to delete. Push it
 * down and put the header where it belongs. Safe to run on every call.
 */
function ensureHeaders(sh) {
  if (sh.getLastRow() === 0) {
    writeHeader(sh);
    return;
  }
  var first = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (String(first[0] || '').trim().toLowerCase() === 'id') return;  // already fine
  sh.insertRowBefore(1);      // rescues the stranded row by moving it to row 2
  writeHeader(sh);
}

function writeHeader(sh) {
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  SpreadsheetApp.flush();
}

function allRows() {
  var sh = sheet();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues();
}

function listSetlists() {
  var seen = {}, order = [];
  allRows().forEach(function (r) {
    var id = String(r[0]);
    if (!id) return;
    if (!seen[id]) {
      seen[id] = { id: id, name: r[1], venue: r[2], date: fmtDate(r[3]),
                   songs: 0, perSet: {}, updated: r[9] ? new Date(r[9]).toISOString() : '' };
      order.push(id);
    }
    seen[id].songs++;
    var sn = Number(r[4]) || 1;
    seen[id].perSet[sn] = (seen[id].perSet[sn] || 0) + 1;
  });
  return order.map(function (id) { return seen[id]; })
              .sort(function (a, b) { return (b.updated || '').localeCompare(a.updated || ''); });
}

function loadSetlist(id) {
  id = String(id || '');
  var meta = null, bucket = {};
  allRows().forEach(function (r) {
    if (String(r[0]) !== id) return;
    if (!meta) meta = { id: id, name: r[1], venue: r[2], date: fmtDate(r[3]) };
    var si = Number(r[4]) || 1;
    (bucket[si] = bucket[si] || []).push({
      pos: Number(r[5]) || 0,
      title: String(r[6] || ''),
      length: fmtLen(r[7]),
      notes: String(r[8] || '')
    });
  });
  if (!meta) throw new Error('Setlist not found');
  var max = Math.max.apply(null, Object.keys(bucket).map(Number).concat([1]));
  meta.sets = [];
  for (var i = 1; i <= max; i++) {
    var arr = bucket[i] || [];
    arr.sort(function (a, b) { return a.pos - b.pos; });
    meta.sets.push(arr.map(function (s) {
      return { title: s.title, length: s.length, notes: s.notes };
    }));
  }
  return meta;
}

function saveSetlist(d) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet();
    var id = String(d.id || '').trim() || ('SL' + Date.now());
    removeRows(sh, id);
    SpreadsheetApp.flush();   // deleteRow leaves getLastRow stale until flushed
    var now = new Date();
    var rows = [];
    (d.sets || []).forEach(function (songs, si) {
      (songs || []).forEach(function (s, pi) {
        rows.push([id, d.name || '', d.venue || '', d.date || '',
                   si + 1, pi + 1, s.title || '', s.length || '', s.notes || '', now]);
      });
    });
    if (!rows.length) throw new Error('Nothing to save — every set is empty');
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    SpreadsheetApp.flush();
    return { id: id, rows: rows.length };
  } finally {
    lock.releaseLock();
  }
}

function deleteSetlist(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return removeRows(sheet(), String(id || ''));
  } finally {
    lock.releaseLock();
  }
}

function removeRows(sh, id) {
  if (sh.getLastRow() < 2) return 0;
  var col = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  var n = 0;
  for (var i = col.length - 1; i >= 0; i--) {
    if (String(col[i][0]) === id) { sh.deleteRow(i + 2); n++; }
  }
  return n;
}

/* Sheets hands back real Date objects for date-ish cells; normalise to text. */
function fmtDate(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'M/d/yyyy');
  }
  return String(v || '');
}

function fmtLen(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'm:ss');
  }
  if (typeof v === 'number' && v > 0 && v < 1) {
    var secs = Math.round(v * 86400);
    return Math.floor(secs / 60) + ':' + ('0' + (secs % 60)).slice(-2);
  }
  return String(v || '');
}
