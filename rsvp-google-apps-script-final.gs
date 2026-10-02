// Pravarthika & Aagash RSVP → Google Sheet
// Target spreadsheet: AP
// Spreadsheet ID: 1KPIMAwC-J98Cwe8LRKyhDXhRS5VteFcMSH1zsRX21xk
// Target tab: Sheet1

const SPREADSHEET_ID = '1KPIMAwC-J98Cwe8LRKyhDXhRS5VteFcMSH1zsRX21xk';
const SHEET_NAME = 'Sheet1';

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: 'Pravarthika & Aagash RSVP' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const familyName = String(payload.familyName || '').trim();
    const totalGuests = Number(payload.totalGuests);

    if (!familyName || !Number.isInteger(totalGuests) || totalGuests < 1 || totalGuests > 15) {
      return jsonResponse({ ok: false, message: 'Invalid RSVP details.' });
    }

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);

    // Add headers only when the sheet is empty.
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Family Name', 'Total Guests']]);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const timestamp = Utilities.formatDate(new Date(), 'Asia/Colombo', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([timestamp, familyName, totalGuests]);

    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, message: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
