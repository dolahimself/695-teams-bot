/**
 * Bear Friends Stories: Google Sheet collector for a pilot.
 *
 * Receives stories from the prototype (STORE=sheet) and adds one row per story to the "Stories" tab.
 * Each story's photos and voice recordings go into a Drive folder for that story, linked from the row.
 *
 * Set up (once):
 *  1. In the team's Google account, create a Google Sheet. Open Extensions > Apps Script and replace
 *     everything in Code.gs with this file. Save.
 *  2. Project Settings (gear icon) > Script properties > Add script property:
 *       SECRET     a long random string. The same value goes in SHEET_SECRET on the server.
 *       FOLDER_ID  optional: the id of a Drive folder for story files (the part after /folders/ in its link).
 *                  Without it, a "Bear Friends Stories files" folder is created in My Drive on the first story.
 *  3. Deploy > New deployment > Select type: Web app. Execute as: Me. Who has access: Anyone.
 *     Deploy, allow access when asked, and copy the Web app URL (it ends in /exec). That is SHEET_URL.
 *  4. If you change this code later: Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy,
 *     so the same URL runs the new code.
 *
 * "Anyone" only means the URL can be called without a Google sign-in. Every request must carry the secret,
 * and the Sheet and the files stay private to this account unless you share them.
 */

const SHEET = 'Stories';
const HEADERS = ['ID', 'Submitted', 'Program', 'What happened', 'What changed', 'Photos', 'Recordings', 'Seconds taken', 'Permission confirmed', 'Status', 'Folder'];
const COL = { id: 1, photos: 6, recordings: 7, folder: 11 };

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return reply({ ok: false, error: 'not JSON' });
  }
  const secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  if (!secret || body.secret !== secret) return reply({ ok: false, error: 'wrong secret' });

  // One request at a time, so two stories arriving together cannot write over each other.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (body.action === 'story') return reply(addStory(body.story || {}));
    if (body.action === 'file') return reply(addFile(body));
    return reply({ ok: false, error: 'unknown action' });
  } catch (err) {
    return reply({ ok: false, error: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

/** Opening the web app URL in a browser shows that it is running. It never shows any stories. */
function doGet() {
  return reply({ ok: true, running: 'Bear Friends Stories collector' });
}

function addStory(s) {
  const sheet = storiesSheet();
  const id = 'S-' + Utilities.getUuid().slice(0, 8); // a letter first, so the Sheet never reads it as a number
  const folder = filesFolder().createFolder(id + ' ' + plain(s.program).slice(0, 60));
  sheet.appendRow([
    id,
    s.submittedAt ? new Date(s.submittedAt) : new Date(),
    text(s.program),
    text(s.whatHappened),
    text(s.whatChanged),
    '',
    '',
    Number(s.durationSeconds) || 0,
    s.permissionConfirmed ? 'Yes' : 'No',
    'Unreviewed',
    folder.getUrl(),
  ]);
  return { ok: true, id: id };
}

function addFile(b) {
  const sheet = storiesSheet();
  const row = findRow(sheet, String(b.storyId || ''));
  if (!row) return { ok: false, error: 'story not found' };
  const folderId = String(sheet.getRange(row, COL.folder).getValue()).match(/folders\/([\w-]+)/);
  if (!folderId) return { ok: false, error: 'story folder not found' };
  const name = String(b.name || 'file').replace(/[^\w.\- ()]+/g, '_').slice(0, 120);
  const blob = Utilities.newBlob(Utilities.base64Decode(String(b.base64 || '')), String(b.type || 'application/octet-stream'), name);
  const file = DriveApp.getFolderById(folderId[1]).createFile(blob);

  const recording = b.kind === 'recording';
  const line = recording ? plain(b.field) + ' (' + mmss(Number(b.seconds) || 0) + '): ' + file.getUrl() : file.getUrl();
  const cell = sheet.getRange(row, recording ? COL.recordings : COL.photos);
  const before = String(cell.getValue() || '');
  cell.setValue(text(before ? before + '\n' + line : line));
  return { ok: true, url: file.getUrl() };
}

function storiesSheet() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = book.getSheetByName(SHEET) || book.insertSheet(SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

function filesFolder() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) return DriveApp.getFolderById(id);
  const folder = DriveApp.createFolder('Bear Friends Stories files');
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

function findRow(sheet, id) {
  if (!id || sheet.getLastRow() < 2) return 0;
  const ids = sheet.getRange(1, COL.id, sheet.getLastRow(), 1).getValues();
  for (let i = ids.length - 1; i >= 1; i--) if (String(ids[i][0]) === id) return i + 1;
  return 0;
}

/** Story text as plain text: a leading = + - or @ would otherwise make the Sheet run it as a formula. */
function text(v) {
  const s = plain(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function plain(v) {
  return v == null ? '' : String(v);
}

function mmss(s) {
  return Math.floor(s / 60) + ':' + ('0' + Math.floor(s % 60)).slice(-2);
}

function reply(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
