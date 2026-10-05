// קוד לגיליון "חתימות תקנון צימרים" - שומר כל חתימה כשורה בגיליון,
// את תמונת החתימה בתיקייה בדרייב, ושולח מייל התראה.
// כל תקנון נכנס ללשונית משלו: צימרים -> "חתימות", כלות -> "כלות".

const NOTIFY_EMAIL = 'info@mettler-winery.com'; // לכאן נשלחת התראה על כל חתימה
const FOLDER_NAME  = 'חתימות תקנון צימרים';
const SHEET_NAME   = 'חתימות';   // לשונית ברירת המחדל (תקנון הצימרים)

const HEADERS = [
  'נחתם ב', 'שם מלא', 'טלפון', 'אימייל', 'הגעה', 'עזיבה',
  'קישור לחתימה', 'גרסת תקנון', 'טביעת אצבע (SHA-256)', 'דפדפן', 'דף', 'מספר אורחים'
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
  const tab = d.doc || SHEET_NAME;
  let sheet = ss.getSheetByName(tab);
  if (!sheet) {
    sheet = ss.insertSheet(tab);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.setRightToLeft(true);
  }

  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);

  const png = Utilities.base64Decode(String(d.signature).split(',')[1] || '');
  const fileName = (d.doc ? d.doc + ' - ' : '') + d.arrival + ' - ' + d.name + '.png';
  const file = folder.createFile(Utilities.newBlob(png, 'image/png', fileName));

  sheet.appendRow([
    new Date(d.signedAt), d.name, "'" + d.phone, d.email, d.arrival, d.departure,
    file.getUrl(), d.termsVersion, d.termsHash, d.userAgent, d.page, d.guests || ''
  ]);

  if (NOTIFY_EMAIL) {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'תקנון ' + (d.doc || 'צימרים') + ' נחתם - ' + d.name + ' (' + d.arrival + ')',
      htmlBody:
        '<div dir="rtl" style="font-family:Arial,sans-serif;font-size:15px">' +
        '<p><b>' + d.name + '</b> חתם/ה על תקנון ' + (d.doc || 'הצימרים') + '.</p>' +
        '<p>טלפון: ' + d.phone + '<br>הגעה: ' + d.arrival + '<br>עזיבה: ' + d.departure +
        (d.guests ? '<br>מספר אורחים: ' + d.guests : '') + '</p>' +
        '<p><img src="cid:sig" style="max-width:300px;border-bottom:1px solid #ccc"></p>' +
        '<p><a href="' + ss.getUrl() + '">לגיליון החתימות</a></p></div>',
      inlineImages: { sig: Utilities.newBlob(png, 'image/png', 'sig.png') }
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
