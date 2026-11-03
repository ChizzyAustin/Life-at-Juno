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
 * Juno Awards: nominations go to the "Award nominations" tab and votes to the
 * "Award votes" tab, one row per room per award per building (sending again
 * updates that row). The "Awards tally" tab counts both. Nominations close at the
 * end of Wednesday 4 November and voting at the end of Thursday 5 November, Lagos time.
 *
 * The tag below limits the script to this one spreadsheet.
 * @OnlyCurrentDoc
 */

const SHEET_NAME = 'Check-ins';
const POINTS_PER_CHECKIN = 1;
const TIMEZONE = 'Africa/Lagos';
const HEADERS = ['Timestamp', 'Event date', 'Event', 'Event ID', 'Building', 'Room', 'Name', 'Points'];

const NOM_SHEET = 'Award nominations';
const NOM_HEADERS = ['Updated', 'Building', 'Award', 'Nominee', 'Nominee key', 'Nominee room', 'Why', 'Nominated by', 'Their room', 'First sent'];
const VOTE_SHEET = 'Award votes';
const VOTE_HEADERS = ['Updated', 'Building', 'Award', 'Vote for', 'Voter', 'Voter room', 'First vote'];
const TALLY_SHEET = 'Awards tally';
const AWARD_IDS = { engaged: 'The Most Engaged Member', neighbour: 'The Ultimate Good Neighbour', trailblazer: 'The Juno Trailblazer Award' };
const BUILDINGS = ['Karimu', 'Emily Akinola'];
const NOMINATIONS_CLOSE = new Date('2026-11-04T23:59:59+01:00');
const VOTING_CLOSE = new Date('2026-11-05T23:59:59+01:00');

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
    if (d.type === 'nominate') return nominate_(d);
    if (d.type === 'vote') return vote_(d);
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

// Juno Awards ---------------------------------------------------------------

function key_(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9 ]/g, '').replace(/\s+/g, ' ').trim();
}

// Writes one row per room, building and award; sending again updates that row.
function upsert_(sh, rows, keyOf, mine, row, firstCol) {
  for (let i = 1; i < rows.length; i++) {
    if (keyOf(rows[i]) === mine) {
      row[firstCol] = rows[i][firstCol] || row[firstCol];
      sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return true;
    }
  }
  sh.appendRow(row);
  rows.push(row);
  return false;
}

function nominate_(d) {
  if (new Date() > NOMINATIONS_CLOSE) return json_({ ok: false, error: 'closed' });
  const building = clean_(d.building, 40);
  const name = clean_(d.name, 80);
  const room = clean_(d.room, 12).toUpperCase();
  const list = Array.isArray(d.nominations) ? d.nominations.slice(0, 3) : [];
  if (BUILDINGS.indexOf(building) < 0 || !name || !room || !list.length) return json_({ ok: false, error: 'missing' });

  const sh = tab_(NOM_SHEET, NOM_HEADERS);
  const rows = sh.getDataRange().getValues();
  const now = new Date();
  let saved = 0;
  list.forEach(function (n) {
    const award = AWARD_IDS[clean_(n.award, 20)];
    const nominee = clean_(n.nominee, 80);
    if (!award || !nominee || key_(nominee) === key_(name)) return;
    const row = [now, building, award, nominee, key_(nominee), clean_(n.nomineeRoom, 12).toUpperCase(),
                 clean_(n.reason, 240), name, room, now];
    upsert_(sh, rows, function (r) { return [r[1], r[2], String(r[8]).toUpperCase()].join('|'); },
            [building, award, room].join('|'), row, 9);
    saved++;
  });
  tally_();
  return json_({ ok: saved > 0, saved: saved, error: saved ? '' : 'missing' });
}

function vote_(d) {
  const now = new Date();
  if (now <= NOMINATIONS_CLOSE || now > VOTING_CLOSE) return json_({ ok: false, error: 'closed' });
  const building = clean_(d.building, 40);
  const name = clean_(d.name, 80);
  const room = clean_(d.room, 12).toUpperCase();
  const list = Array.isArray(d.votes) ? d.votes.slice(0, 3) : [];
  if (BUILDINGS.indexOf(building) < 0 || !name || !room || !list.length) return json_({ ok: false, error: 'missing' });

  const sh = tab_(VOTE_SHEET, VOTE_HEADERS);
  const rows = sh.getDataRange().getValues();
  let saved = 0;
  list.forEach(function (v) {
    const award = AWARD_IDS[clean_(v.award, 20)];
    const nominee = clean_(v.nominee, 80);
    if (!award || !nominee) return;
    const row = [now, building, award, nominee, name, room, now];
    upsert_(sh, rows, function (r) { return [r[1], r[2], String(r[5]).toUpperCase()].join('|'); },
            [building, award, room].join('|'), row, 6);
    saved++;
  });
  tally_();
  return json_({ ok: saved > 0, saved: saved, error: saved ? '' : 'missing' });
}

// Builds the "Awards tally" tab once: live counts of nominations (by name) and votes.
function tally_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(TALLY_SHEET)) return;
  tab_(NOM_SHEET, NOM_HEADERS);
  tab_(VOTE_SHEET, VOTE_HEADERS);
  const sh = ss.insertSheet(TALLY_SHEET);
  sh.getRange('A1').setValue('Nominations (by building, award and name)').setFontWeight('bold');
  sh.getRange('A2').setFormula("=IFERROR(QUERY('" + NOM_SHEET + "'!B:E, \"select B, C, E, count(D) where B is not null group by B, C, E order by B, C, count(D) desc label E 'Nominee', count(D) 'Nominations'\", 1), \"No nominations yet\")");
  sh.getRange('G1').setValue('Votes (by building, award and name)').setFontWeight('bold');
  sh.getRange('G2').setFormula("=IFERROR(QUERY('" + VOTE_SHEET + "'!B:D, \"select B, C, D, count(D) where B is not null group by B, C, D order by B, C, count(D) desc label count(D) 'Votes'\", 1), \"No votes yet\")");
  sh.setFrozenRows(2);
}

// The web app only accepts check-ins, RSVPs, award nominations and votes. Room totals and guest lists stay
// inside the Sheet (and AppSheet, if you add it), so room numbers are never public.
function doGet() {
  return json_({ ok: true, service: 'Life at Juno check-in, RSVP and awards' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
