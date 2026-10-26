/**
 * Receives questions from the website form and saves them to this Google Sheet.
 * Setup steps are in README.md ("Connect the question form").
 */
const SHEET_NAME   = 'Questions';
const NOTIFY_EMAIL = 'maheshdattaastro@gmail.com';   // email to alert on each new question; leave '' for no email

const HEADERS = ['Received', 'Name', 'Email', 'WhatsApp / phone', 'Topic',
                 'Date of birth', 'Time of birth', 'Place of birth', 'Question',
                 'Status', 'Reply notes'];

function doPost(e) {
  const p = (e && e.parameter) ? e.parameter : {};

  // Spam trap: real visitors leave this hidden field empty
  if (p._gotcha) return ContentService.createTextOutput('ok');
  if (!p.name || !p.email || !p.question) return ContentService.createTextOutput('missing fields');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }

    const row = [
      Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMM yyyy, hh:mm a'),
      p.name, p.email, p.phone || '', p.topic || '',
      p.date_of_birth || '', p.time_of_birth || '', p.place_of_birth || '',
      p.question, 'New', ''
    ];

    // Save as plain text so phone numbers like +91... and any "=" text are kept as typed
    const r = sheet.getLastRow() + 1;
    const range = sheet.getRange(r, 1, 1, row.length);
    range.setNumberFormat('@');
    range.setValues([row]);
    range.setWrap(true).setVerticalAlignment('top');
  } finally {
    lock.releaseLock();
  }

  if (NOTIFY_EMAIL) {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      replyTo: p.email,
      subject: 'New astrology question from ' + p.name,
      body: 'Topic: ' + (p.topic || '-') + '\n' +
            'Phone: ' + (p.phone || '-') + '\n' +
            'Birth: ' + (p.date_of_birth || '-') + ' ' + (p.time_of_birth || '') + ', ' + (p.place_of_birth || '-') + '\n\n' +
            p.question + '\n\nReply to this email to answer ' + p.name + '.'
    });
  }
  return ContentService.createTextOutput('ok');
}

// Opening the web app URL in a browser shows this, so you can check it is live
function doGet() {
  return ContentService.createTextOutput('The question form is connected.');
}
