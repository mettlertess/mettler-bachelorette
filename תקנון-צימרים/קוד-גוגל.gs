// קוד לגיליון "חתימות תקנון צימרים" - שומר כל חתימה כשורה בגיליון,
// את תמונת החתימה בתיקייה בדרייב, ושולח מייל התראה.

const NOTIFY_EMAIL = Session.getEffectiveUser().getEmail(); // התראה לבעל החשבון. ריק = בלי מייל
const FOLDER_NAME  = 'חתימות תקנון צימרים';
const SHEET_NAME   = 'חתימות';

const HEADERS = [
  'נחתם ב', 'שם מלא', 'טלפון', 'אימייל', 'הגעה', 'עזיבה',
  'קישור לחתימה', 'גרסת תקנון', 'טביעת אצבע (SHA-256)', 'דפדפן', 'דף'
];

// פתיחת הקישור בדפדפן - לבדיקה שהחיבור חי
function doGet() {
  return ContentService.createTextOutput('החיבור לתקנון הצימרים פעיל');
}

// הרצה ידנית אחת לאישור הרשאות (דרייב, גיליונות, מייל)
function authorize() {
  DriveApp.getRootFolder();
  SpreadsheetApp.getActiveSpreadsheet();
  MailApp.getRemainingDailyQuota();
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.setRightToLeft(true);
  }

  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);

  const png = Utilities.base64Decode(String(d.signature).split(',')[1] || '');
  const fileName = d.arrival + ' - ' + d.name + '.png';
  const file = folder.createFile(Utilities.newBlob(png, 'image/png', fileName));

  sheet.appendRow([
    new Date(d.signedAt), d.name, "'" + d.phone, d.email, d.arrival, d.departure,
    file.getUrl(), d.termsVersion, d.termsHash, d.userAgent, d.page
  ]);

  if (NOTIFY_EMAIL) {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'תקנון צימרים נחתם - ' + d.name + ' (' + d.arrival + ')',
      htmlBody:
        '<div dir="rtl" style="font-family:Arial,sans-serif;font-size:15px">' +
        '<p><b>' + d.name + '</b> חתם/ה על תקנון האירוח.</p>' +
        '<p>טלפון: ' + d.phone + '<br>הגעה: ' + d.arrival + '<br>עזיבה: ' + d.departure + '</p>' +
        '<p><img src="cid:sig" style="max-width:300px;border-bottom:1px solid #ccc"></p>' +
        '<p><a href="' + ss.getUrl() + '">לגיליון החתימות</a></p></div>',
      inlineImages: { sig: Utilities.newBlob(png, 'image/png', 'sig.png') }
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
