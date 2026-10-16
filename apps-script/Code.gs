/**
 * Life at Juno: event check-ins.
 *
 * Paste this into Extensions > Apps Script on the "Juno Check-ins" Google Sheet,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Copy the web app URL into CHECKIN_URL at the top of checkin/index.html.
 *
 * Each check-in adds one row to the "Check-ins" tab. A person can check in to
 * an event once, and only on the day of the event (Lagos time).
 */

const SHEET_NAME = 'Check-ins';
const POINTS_PER_CHECKIN = 1;
const TIMEZONE = 'Africa/Lagos';
const HEADERS = ['Timestamp', 'Event date', 'Event', 'Event ID', 'Building', 'Room', 'Name', 'Points'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function clean_(value, max) {
  return String(value == null ? '' : value).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const name = clean_(d.name, 60);
    const building = clean_(d.building, 40);
    const room = clean_(d.room, 12).toUpperCase();
    const eventId = clean_(d.eventId, 100);
    const eventTitle = clean_(d.eventTitle, 120);
    const eventDate = clean_(d.eventDate, 10);

    if (!name || !building || !room || !eventId) return json_({ ok: false, error: 'missing' });

    const today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
    if (eventDate !== today) return json_({ ok: false, error: 'not-today' });

    const sh = sheet_();
    const rows = sh.getDataRange().getValues();
    const already = rows.slice(1).some(function (r) {
      return r[3] === eventId && String(r[5]).toUpperCase() === room &&
             String(r[6]).toLowerCase() === name.toLowerCase();
    });
    if (!already) {
      sh.appendRow([new Date(), eventDate, eventTitle, eventId, building, room, name, POINTS_PER_CHECKIN]);
    }
    const all = already ? rows : sh.getDataRange().getValues();
    return json_({ ok: true, duplicate: already, roomPoints: roomPoints_(all, building, room) });
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}

function roomPoints_(rows, building, room) {
  let points = 0;
  rows.slice(1).forEach(function (r) {
    if (r[4] === building && String(r[5]).toUpperCase() === room) points += Number(r[7]) || 0;
  });
  return points;
}

// The web app only accepts check-ins. Room totals stay inside the Sheet
// (and AppSheet, if you add it), so room numbers are never public.
function doGet() {
  return json_({ ok: true, service: 'Life at Juno check-in' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
