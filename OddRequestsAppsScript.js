const SPREADSHEET_ID = '1KmpS3C6H8M7b5s1yOeVe42dzJTCuC2Djn5aKToALwkY';
const SHEET_NAME = 'OddRequests';

function doPost(e) {
  const sheet = getOddRequestsSheet_();
  const params = e && e.parameter ? e.parameter : {};
  const name = String(params.name || '').trim();
  const submittedAt = String(params.submittedAt || '').trim() || new Date().toISOString();
  const request = String(params.request || '').trim();

  if (!name || !request) {
    return json_({ ok: false, error: 'Missing name or request.' });
  }

  sheet.appendRow([name, submittedAt, request]);
  return json_({ ok: true });
}

function getOddRequestsSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Name', 'Date/Time of Request', 'Request']);
  }

  return sheet;
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
