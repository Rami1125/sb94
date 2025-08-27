const SPREADSHEET_ID = '1TPwAP0h05IyzvusybJv3zMKSOKpFNUS9jZtEK5pSgps'; 
const SHEET_NAME = 'מעקב'; // Ensure this matches your sheet name
const OVERDUE_THRESHOLD_DAYS = 10;
const DEFAULT_RECIPIENT_EMAILS = "ramims@saban94.co.il,rami.msarwa1@gmail.com";
function doGet(e) {
  try {
    const action = e.parameter.action;
    let responseData = { success: false, message: 'פעולה לא ידועה.' };

    if (action === 'sendDailyReport') {
      const recipientEmails = e.parameter.recipientEmails; // קבל את כתובות המייל מהבקשה
      sendDailyReport(recipientEmails);
      responseData = { success: true, message: "דוח יומי נשלח בהצלחה!" };
    } else {
      // אם יש לכם פעולות נוספות ל-Apps Script, תוסיפו אותן כאן
      responseData.message = `פעולה "${action}" לא נתמכת בסקריפט זה.`;
    }

    return ContentService.createTextOutput(JSON.stringify(responseData))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log("שגיאה ב-doGet: " + error.message);
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "שגיאה בביצוע הפעולה: " + error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * מכין ושולח דוח יומי (דמו) לכתובות מייל מרובות.
 * @param {string} recipientEmailsString מחרוזת של כתובות מייל מופרדות בפסיקים (לדוגמה: "a@b.com,c@d.com").
 */
function sendDailyReport(recipientEmailsString) {
  if (!recipientEmailsString) {
    throw new Error("לא סופקו כתובות מייל לנמענים.");
  }

  Logger.log("מכין דוח יומי עבור: " + recipientEmailsString);

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error(`גיליון בשם "${SHEET_NAME}" לא נמצא בגיליון העבודה.`);
  }

  const data = sheet.getDataRange().getDisplayValues(); // קבל את כל הנתונים מהגיליון
  
  // יצירת תוכן ה-HTML של הדוח
  let reportHtml = `
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; direction: rtl; text-align: right; }
        h1 { color: #2E8B57; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: right; }
        th { background-color: #f2f2f2; }
        .status-open { color: green; font-weight: bold; }
        .status-overdue { color: red; font-weight: bold; }
        .status-closed { color: gray; }
      </style>
    </head>
    <body>
      <h1>דוח יומי - מכולות</h1>
      <p>שלום,</p>
      <p>זהו עדכון הדוח היומי שלך מיום ${new Date().toLocaleDateString('he-IL')}.</p>
      <table>
  `;

  if (data.length > 0) {
    // הוספת כותרות עמודות
    reportHtml += `<thead><tr>`;
    data[0].forEach(header => reportHtml += `<th>${header}</th>`);
    reportHtml += `</tr></thead><tbody>`;

    // הוספת שורות נתונים (מדלג על שורת הכותרות)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // כאן ניתן להוסיף לוגיקה לעיבוד או סגנון מותנה, לדוגמה:
      // נניח שהסטטוס הוא בעמודה ה-8 (אינדקס 7)
      const status = row[7] || ''; // 'פתוח', 'חורג', 'סגור'
      let statusClass = '';
      if (status === 'פתוח') statusClass = 'status-open';
      else if (status === 'חורג') statusClass = 'status-overdue';
      else if (status === 'סגור') statusClass = 'status-closed';

      reportHtml += `<tr>`;
      row.forEach((cell, cellIndex) => {
        if (cellIndex === 7) { // עמודת סטטוס
          reportHtml += `<td class="${statusClass}">${cell}</td>`;
        } else {
          reportHtml += `<td>${cell}</td>`;
        }
      });
      reportHtml += `</tr>`;
    }
    reportHtml += `</tbody>`;
  } else {
    reportHtml += `<tbody><tr><td colspan="100%">אין נתונים זמינים לדוח.</td></tr></tbody>`;
  }

  reportHtml += `
      </table>
      <p>בברכה,</p>
      <p>צוות המכולות</p>
    </body>
    </html>
  `;

  const subject = "דוח יומי - מערכת CRM מכולות";

  try {
    // MailApp.sendEmail תומכת במחרוזת של מיילים מופרדים בפסיקים בשדה 'to'
    MailApp.sendEmail({
      to: recipientEmailsString,
      subject: subject,
      htmlBody: reportHtml // שליחת הדוח כ-HTML
    });
    Logger.log("דוח נשלח בהצלחה ל: " + recipientEmailsString);
  } catch (e) {
    Logger.log("שגיאה בשליחת המייל: " + e.toString());
    throw new Error("שגיאה בשליחת המייל: " + e.message);
  }
}

/**
 * פונקציה זו יכולה לשמש כטריגר מבוסס זמן.
 * היא תשלח את הדוח היומי באופן אוטומטי לכתובות המייל הקבועות.
 * 🚨 חשוב: עדכן את רשימת הנמענים כאן אם אתה משתמש בטריגר מבוסס זמן.
 */
function sendDailyReportViaTrigger() {
  const defaultRecipients = "ramims@saban94.co.il,rami.msarwa1@gmail.com"; // 🚨 עדכן כאן את כתובות המייל הקבועות
  Logger.log("טריגר יומי מופעל. שולח דוח ל: " + defaultRecipients);
  sendDailyReport(defaultRecipients);
}
