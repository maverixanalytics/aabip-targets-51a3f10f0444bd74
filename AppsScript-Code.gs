/**
 * AABIP 2026 — target-list sync + badge-photo lead capture
 *
 * Replaces the existing Code.gs in your Apps Script project.
 * Everything the ATS/AATS/DDW pages already do is unchanged — this only ADDS
 * badge-photo capture and the RSVP flag. Existing rows are untouched.
 *
 * SETUP
 *  1. Extensions -> Apps Script, select all of Code.gs, paste this over it.
 *  2. Check SHEET_ID below still matches your sheet (it should already).
 *  3. Deploy -> Manage deployments -> pencil icon -> Version: New version -> Deploy.
 *     Use MANAGE deployments, not "New deployment" — that would change the URL
 *     and the ATS/AATS/DDW pages would stop syncing.
 *  4. First run asks for Drive permission (it needs to create the photo folder).
 */

const SHEET_ID     = 'PASTE_SHEET_ID_HERE';          // unchanged from your current script
const STATE_SHEET  = 'Sheet1';                       // id | state | timestamp
const LEAD_SHEET   = 'Leads';                        // created automatically
const PHOTO_FOLDER = 'AABIP 2026 Badge Photos';      // created automatically in My Drive

function ss_()          { return SpreadsheetApp.openById(SHEET_ID); }
function stateSheet_()  { return ss_().getSheetByName(STATE_SHEET); }

function leadSheet_() {
  let sh = ss_().getSheetByName(LEAD_SHEET);
  if (!sh) {
    sh = ss_().insertSheet(LEAD_SHEET);
    sh.appendRow(['Timestamp', 'Physician ID', 'Physician', 'Rep', 'Note', 'Photo', 'Points', 'UID']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function photoFolder_() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}

/* ------------------------------------------------------------------ read */

function doGet(e) {
  const data = stateSheet_().getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) result[String(data[i][0])] = true;
  }

  // Photo counts per physician, so a card can show "2 captured".
  const lv = leadSheet_().getDataRange().getValues();
  const counts = {};
  for (let i = 1; i < lv.length; i++) {
    const id = String(lv[i][1] || '');
    if (id) counts[id] = (counts[id] || 0) + 1;
  }
  Object.keys(counts).forEach(function (k) { result[k + '-photos'] = counts[k]; });

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ----------------------------------------------------------------- write */

function doPost(e) {
  const p = JSON.parse(e.postData.contents);
  return (p.kind === 'lead') ? saveLead_(p) : saveState_(p);
}

// Checkbox state: {id}-conf, {id}-email, {id}-rsvp. Absence of a row means false.
function saveState_(params) {
  const sheet = stateSheet_();
  const id = String(params.id);
  const on = !!params.connected;
  const data = sheet.getDataRange().getValues();

  let row = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) { row = i + 1; break; }
  }
  if (on) {
    if (row === -1) sheet.appendRow([id, 'connected', new Date().toISOString()]);
    else {
      sheet.getRange(row, 2).setValue('connected');
      sheet.getRange(row, 3).setValue(new Date().toISOString());
    }
  } else if (row !== -1) {
    sheet.deleteRow(row);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Badge photo + note from a rep's phone.
function saveLead_(p) {
  const sheet = leadSheet_();

  // A flaky hall connection can lose the response to a write that actually
  // succeeded; the phone then retries. Reject the replay instead of duplicating.
  if (p.uid && sheet.getLastRow() > 1) {
    const seen = sheet.getRange(2, 8, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < seen.length; i++) {
      if (String(seen[i][0]) === String(p.uid)) {
        return ContentService.createTextOutput(JSON.stringify({ ok: true, duplicate: true }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
  }

  let url = '';
  if (p.photo) {
    const bytes = Utilities.base64Decode(p.photo);
    const name  = (p.id || 'lead') + '_' + (p.rep || 'rep').replace(/[^A-Za-z0-9]/g, '') +
                  '_' + new Date().getTime() + '.jpg';
    const file  = photoFolder_().createFile(Utilities.newBlob(bytes, p.mime || 'image/jpeg', name));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    url = file.getUrl();
  }

  // Contest scoring: a photo WITH a substantive note counts double.
  const note   = String(p.note || '').trim();
  const points = (url && note.length >= 15) ? 2 : 1;

  sheet.appendRow([
    new Date().toISOString(), p.id || '', p.name || '', p.rep || '', note, url, points, p.uid || ''
  ]);

  return ContentService.createTextOutput(JSON.stringify({ ok: true, url: url, points: points }))
    .setMimeType(ContentService.MimeType.JSON);
}
