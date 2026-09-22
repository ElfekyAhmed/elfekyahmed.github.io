/**
 * Al Asriyah — Marmomac kiosk → Google Sheet
 *
 * 1. Open your Google Sheet → Extensions → Apps Script.
 * 2. Delete whatever is in Code.gs, paste this whole file, press Save (⌘S).
 * 3. Deploy → New deployment → type "Web app"
 *      Execute as:      Me
 *      Who has access:  Anyone          ← required, the kiosk is not signed in to Google
 *    → Deploy → authorise when asked → copy the Web app URL (ends in /exec).
 * 4. On the kiosk: tap the triangle mark 5×, PIN 2031, paste the URL in
 *    "Google Sheet endpoint", Save, then Test.
 *
 * If you later edit this script, use Deploy → Manage deployments → ✎ → Version: New
 * so the same URL picks up the change.
 */

const SHEET_NAME = 'Leads';
const HEADERS = ['ID', 'Name', 'Phone', 'Email', 'Country', 'Country code', 'Signed up at', 'Source', 'Received at'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (data.action !== 'lead' || !data.id || !data.name) {
      return out({ ok: false, error: 'Missing lead fields' });
    }
    const sheet = getSheet();
    if (findId(sheet, data.id)) return out({ ok: true, duplicate: true });

    const row = sheet.getLastRow() + 1;
    const range = sheet.getRange(row, 1, 1, HEADERS.length);
    range.setNumberFormat('@'); // keep "+968 …" as text, not a formula
    range.setValues([[
      String(data.id),
      String(data.name || ''),
      String(data.phone || ''),
      String(data.email || ''),
      String(data.country || ''),
      String(data.countryCode || ''),
      String(data.createdAt || ''),
      String(data.source || ''),
      new Date().toISOString()
    ]]);
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Used by the kiosk's "Test" button (and handy to open in a browser).
function doGet() {
  try {
    const sheet = getSheet();
    return out({ ok: true, leads: Math.max(sheet.getLastRow() - 1, 0), sheet: SHEET_NAME });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function findId(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return false;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return true;
  return false;
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
