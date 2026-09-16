const SPREADSHEET_ID = '1KmpS3C6H8M7b5s1yOeVe42dzJTCuC2Djn5aKToALwkY';
const SHEET_NAME = 'Riley Lab Portal REQUESTS';
const SLACK_WEBHOOK_PROPERTY = 'SLACK_WEBHOOK_URL';

function doPost(e) {
  const sheet = getOddRequestsSheet_();
  const params = e && e.parameter ? e.parameter : {};
  const name = String(params.name || '').trim();
  const submittedAt = String(params.submittedAt || '').trim() || new Date().toISOString();
  const request = String(params.request || '').trim();

  if (!name || !request) {
    return json_({ ok: false, error: 'Missing name or request.' });
  }

  sheet.appendRow([submittedAt, name, request]);
  notifySlack_(submittedAt, name, request);
  return json_({ ok: true });
}

function getOddRequestsSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date', 'Name', 'Request']);
  }

  return sheet;
}

function notifySlack_(submittedAt, name, request) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty(SLACK_WEBHOOK_PROPERTY);
  if (!webhookUrl) {
    console.warn(`Missing script property: ${SLACK_WEBHOOK_PROPERTY}`);
    return;
  }

  const message = [
    ':bell: *New Riley Lab Portal request*',
    `*Date:* ${submittedAt}`,
    `*Name:* ${name}`,
    `*Request:* ${request}`
  ].join('\n');

  const response = UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: message }),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code < 200 || code >= 300) {
    console.error(`Slack notification failed: ${code} ${text}`);
  }
}

function testSlackNotification() {
  notifySlack_(new Date().toISOString(), 'Test User', 'This is a test from Riley Lab Portal Apps Script.');
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
