/**
 * Life at Juno: event check-ins and RSVPs.
 *
 * Paste this into Extensions > Apps Script on the "Juno Check-ins" Google Sheet,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Copy the web app URL into CHECKIN_URL in checkin/index.html and RSVP_URL in
 * rsvp/index.html. When you change this code later, use Deploy > Manage
 * deployments > Edit > Version: New version, so the URL stays the same.
 *
 * Check-ins: each one adds a row to the "Check-ins" tab. A person can check in
 * to an event once, and only on the day of the event (Lagos time).
 *
 * RSVPs: each reply adds a row to the "RSVPs" tab. If the same person replies
 * again for the same event, their row is updated rather than duplicated.
 *
 * The tag below limits the script to this one spreadsheet.
 * @OnlyCurrentDoc
 */

const SHEET_NAME = 'Check-ins';
const POINTS_PER_CHECKIN = 1;
const TIMEZONE = 'Africa/Lagos';
const HEADERS = ['Timestamp', 'Event date', 'Event', 'Event ID', 'Building', 'Room', 'Name', 'Points'];

const RSVP_SHEET = 'RSVPs';
const RSVP_HEADERS = ['Updated', 'Event ID', 'Event', 'Response', 'Name', 'Guest type',
                      'Building', 'Room', 'Organisation', 'Phone', 'Dietary needs', 'First reply'];

function tab_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function sheet_() {
  return tab_(SHEET_NAME, HEADERS);
}

function clean_(value, max) {
  return String(value == null ? '' : value).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.type === 'rsvp') return rsvp_(d);
    return checkin_(d);
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}

function checkin_(d) {
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
}

function roomPoints_(rows, building, room) {
  let points = 0;
  rows.slice(1).forEach(function (r) {
    if (r[4] === building && String(r[5]).toUpperCase() === room) points += Number(r[7]) || 0;
  });
  return points;
}

function rsvp_(d) {
  const eventId = clean_(d.eventId, 100);
  const eventTitle = clean_(d.eventTitle, 120);
  const response = clean_(d.response, 5).toLowerCase();
  const name = clean_(d.name, 80);
  const guestType = clean_(d.guestType, 10) === 'guest' ? 'guest' : 'member';
  const building = guestType === 'member' ? clean_(d.building, 40) : '';
  const room = guestType === 'member' ? clean_(d.room, 12).toUpperCase() : '';
  const org = guestType === 'guest' ? clean_(d.organisation, 80) : '';
  const phone = guestType === 'guest' ? clean_(d.phone, 24) : '';
  const dietary = response === 'yes' ? clean_(d.dietary, 140) : '';

  if (!eventId || !name || (response !== 'yes' && response !== 'no')) return json_({ ok: false, error: 'missing' });
  if (guestType === 'member' && (!building || !room)) return json_({ ok: false, error: 'missing' });

  const sh = tab_(RSVP_SHEET, RSVP_HEADERS);
  const rows = sh.getDataRange().getValues();
  const key = function (r) {
    return [r[1], String(r[4]).toLowerCase(), String(r[5]).toLowerCase(), String(r[7]).toUpperCase(), String(r[8]).toLowerCase()].join('|');
  };
  const mine = [eventId, name.toLowerCase(), guestType, room, org.toLowerCase()].join('|');
  let at = -1;
  for (let i = 1; i < rows.length; i++) { if (key(rows[i]) === mine) { at = i; break; } }

  const now = new Date();
  const first = at > 0 ? rows[at][11] : now;
  const row = [now, eventId, eventTitle, response === 'yes' ? 'Yes' : 'No', name, guestType === 'guest' ? 'Guest' : 'Member',
               building, room, org, phone, dietary, first];
  if (at > 0) sh.getRange(at + 1, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return json_({ ok: true, updated: at > 0 });
}

// The web app only accepts check-ins and RSVPs. Room totals and guest lists stay
// inside the Sheet (and AppSheet, if you add it), so room numbers are never public.
function doGet() {
  return json_({ ok: true, service: 'Life at Juno check-in and RSVP' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
