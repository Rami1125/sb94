/**
 * קונפיגורציה של הקובץ.
 * יש לעדכן את המזהה ואת שם הגיליון בהתאם לקובץ שלכם.
 */
const SPREADSHEET_ID = '1TPwAP0h05IyzvusybJv3zMKSOKpFNUS9jZtEK5pSgps';
const SHEET_NAME = 'מעקב';
const DEFAULT_RECIPIENT_EMAILS = "ramims@saban94.co.il,rami.msarwa1@gmail.com";
const OVERDUE_THRESHOLD_DAYS = 10;
const EMAIL_SUBJECT = "דוח יומי - מערכת CRM מכולות";

/**
 * פונקציה ראשית שמכינה ושולחת את הדוח היומי.
 * היא יכולה להיות מופעלת על ידי טריגר מבוסס זמן.
 */
function sendDailyReport() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log(`Error: Sheet named "${SHEET_NAME}" not found.`);
    return;
  }

  const data = sheet.getDataRange().getDisplayValues();

  if (data.length <= 1) {
    Logger.log("No data available to create a report.");
    return;
  }

  const headers = data[0];
  const allOrders = data.slice(1).map(row => {
    let order = {};
    headers.forEach((header, index) => {
      order[header] = row[index];
    });
    return order;
  });

  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const yesterdayString = Utilities.formatDate(yesterday, "GMT+2", "dd/MM/yyyy");

  // פילוח הנתונים לפי סטאטוס
  const openOrders = allOrders.filter(order => order['סטטוס'] === 'פתוח' && order['לקוח'] && order['לקוח'].trim() !== '');
  const overdueOrders = allOrders.filter(order => order['סטטוס'] === 'חורג' && order['לקוח'] && order['לקוח'].trim() !== '');
  const newOrders = allOrders.filter(order => {
    const orderDate = order['תאריך הזמנה'];
    return orderDate && Utilities.formatDate(new Date(orderDate), "GMT+2", "dd/MM/yyyy") === yesterdayString;
  });

  // חישובים
  const totalOpenContainers = openOrders.reduce((sum, order) => sum + parseInt(order['מספר מכולות'] || 0), 0);
  const totalOverdueContainers = overdueOrders.reduce((sum, order) => sum + parseInt(order['מספר מכולות'] || 0), 0);
  const totalUsedContainers = totalOpenContainers + totalOverdueContainers;

  // יצירת תוכן ה-HTML של הדוח
  const htmlBody = generateReportHtml(
    totalUsedContainers,
    totalOpenContainers,
    totalOverdueContainers,
    openOrders,
    overdueOrders,
    newOrders
  );

  try {
    MailApp.sendEmail({
      to: DEFAULT_RECIPIENT_EMAILS,
      subject: EMAIL_SUBJECT,
      htmlBody: htmlBody
    });
    Logger.log("Daily report sent successfully to: " + DEFAULT_RECIPIENT_EMAILS);
  } catch (e) {
    Logger.log("Error sending email: " + e.toString());
  }
}

/**
 * פונקציה ליצירת טבלת HTML מנתונים.
 * @param {Array<Object>} orders - מערך של אובייקטי הזמנות.
 * @param {string} title - כותרת הטבלה.
 * @param {string} className - שם המחלקה לעיצוב.
 * @returns {string} - מחרוזת HTML של הטבלה.
 */
function createHtmlTable(orders, title, className) {
  if (orders.length === 0) {
    return `<div class="info-box ${className}"><h3>${title}</h3><p style="text-align: center;">אין נתונים זמינים.</p></div>`;
  }

  // סינון כותרות רלוונטיות
  const headers = ['תאריך הזמנה', 'תעודה', 'שם סוכן', 'שם לקוח', 'כתובת', 'סוג פעולה', 'ימים שעברו', 'מספר מכולות', 'סטטוס', 'תאריך סיום צפוי'];

  let tableHtml = `<div class="table-container ${className}">`;
  tableHtml += `<h3 style="text-align: center;">${title}</h3>`;
  tableHtml += `<table cellpadding="0" cellspacing="0" style="width: 100%; border-collapse: collapse; margin-top: 10px;"><thead><tr>`;

  headers.forEach(header => {
    tableHtml += `<th style="border: 1px solid #ddd; padding: 8px; text-align: right; background-color: #f2f2f2;">${header}</th>`;
  });

  tableHtml += `</tr></thead><tbody>`;

  orders.forEach(order => {
    tableHtml += `<tr>`;
    headers.forEach(header => {
      let cellValue = order[header] || '';
      let style = 'border: 1px solid #ddd; padding: 8px; text-align: right;';
      
      // עיצוב תאריך
      if (header.includes('תאריך')) {
        try {
          cellValue = Utilities.formatDate(new Date(cellValue), "GMT+2", "dd/MM/yyyy");
        } catch(e) {
          cellValue = order[header]; // שימור הערך המקורי אם יש שגיאת המרה
        }
      }

      // הדגשת סטטוס
      if (header === 'סטטוס') {
        if (cellValue === 'פתוח') style += ' font-weight: bold; color: green;';
        else if (cellValue === 'חורג') style += ' font-weight: bold; color: red;';
        else if (cellValue === 'סגור') style += ' color: gray;';
      }

      tableHtml += `<td style="${style}">${cellValue}</td>`;
    });
    tableHtml += `</tr>`;
  });

  tableHtml += `</tbody></table></div>`;
  return tableHtml;
}

/**
 * פונקציה שמרכזת את יצירת כל גוף המייל כ-HTML.
 * @param {number} totalUsedContainers
 * @param {number} totalOpenContainers
 * @param {number} totalOverdueContainers
 * @param {Array<Object>} openOrders
 * @param {Array<Object>} overdueOrders
 * @param {Array<Object>} newOrders
 * @returns {string} - מחרוזת HTML מלאה.
 */
function generateReportHtml(totalUsedContainers, totalOpenContainers, totalOverdueContainers, openOrders, overdueOrders, newOrders) {
  const todayDate = new Date().toLocaleDateString('he-IL', { day: 'numeric', month: 'numeric', year: 'numeric' });
  
  // שימוש בפונקציה createHtmlTable
  const newOrdersTable = createHtmlTable(newOrders, 'הזמנות/פעילויות מהיום הקודם', 'new-orders-table');
  const openTable = createHtmlTable(openOrders, 'לקוחות פעילים', 'open-table');
  const overdueTable = createHtmlTable(overdueOrders, 'לקוחות חורגים', 'overdue-table');

  return `
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        body { 
          font-family: Arial, sans-serif; 
          direction: rtl; 
          text-align: right; 
          background-color: #f0f0f0; 
          margin: 0; 
          padding: 20px;
        }
        .container {
          max-width: 800px;
          margin: 0 auto;
          background-color: #ffffff;
          border-radius: 10px;
          box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
          padding: 20px;
        }
        h1, h2, h3 { 
          color: #1a73e8; 
          border-bottom: 2px solid #1a73e8; 
          padding-bottom: 5px;
          text-align: center;
        }
        .summary-box {
          background-color: #e8f0fe;
          border: 1px solid #1a73e8;
          border-radius: 8px;
          padding: 15px;
          margin-bottom: 20px;
          text-align: center;
          font-weight: bold;
          color: #1a73e8;
        }
        .summary-flex {
            display: flex;
            flex-wrap: wrap;
            justify-content: space-around;
        }
        .summary-item {
            background-color: #d1e2ff;
            border-radius: 5px;
            padding: 10px;
            margin: 5px;
            min-width: 200px;
            text-align: center;
        }
        table { 
          width: 100%; 
          border-collapse: collapse; 
          margin-top: 20px; 
          table-layout: fixed; /* שינוי כדי למנוע טבלאות חורגות */
        }
        th, td { 
          border: 1px solid #ddd; 
          padding: 8px; 
          text-align: right; 
          word-wrap: break-word; /* מונע גלישת טקסט מחוץ לתא */
        }
        th { 
          background-color: #f2f2f2; 
          color: #333;
        }
        .open-table h3 { color: #1a73e8; }
        .open-table th { background-color: #c7e0ff; }
        
        .overdue-table h3 { color: #d93025; }
        .overdue-table th { background-color: #f4cccc; }
        .overdue-table { background-color: #fff8f8; } /* רקע אדום עדין */

        .new-orders-table h3 { color: #1e8e3e; }
        .new-orders-table th { background-color: #d7f5e1; }

        .table-container {
          border: 1px solid #ccc;
          border-radius: 8px;
          padding: 15px;
          margin-bottom: 20px;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>דוח יומי למערכת ניהול מכולות</h1>
        <p style="text-align: center;">שלום,</p>
        <p style="text-align: center;">זהו דוח עדכון יומי מקיף מיום <b>${todayDate}</b>. הדוח כולל נתוני מכולות בשימוש, הזמנות פתוחות וחורגות, וכן סיכום הזמנות חדשות.</p>
        
        <h2>סיכום נתונים</h2>
        <div class="summary-box">
            <div class="summary-flex">
                <div class="summary-item">
                    <p>סה"כ מכולות בשימוש:</p>
                    <p style="font-size: 24px; color: #1a73e8;"><b>${totalUsedContainers}</b></p>
                </div>
                <div class="summary-item">
                    <p>מכולות בהזמנות פתוחות:</p>
                    <p style="font-size: 24px; color: #1e8e3e;"><b>${totalOpenContainers}</b></p>
                </div>
                <div class="summary-item">
                    <p>מכולות בהזמנות חורגות:</p>
                    <p style="font-size: 24px; color: #d93025;"><b>${totalOverdueContainers}</b></p>
                </div>
            </div>
        </div>

        ${newOrdersTable}
        ${openTable}
        ${overdueTable}

        <p style="text-align: center;">בברכה,</p>
        <p style="text-align: center;">צוות המכולות</p>
      </div>
    </body>
    </html>
  `;
}

/**
 * פונקציה זו יכולה לשמש כטריגר מבוסס זמן.
 * היא מפעילה את הפונקציה הראשית לשליחת הדוח.
 */
function sendDailyReportViaTrigger() {
  Logger.log("טריגר יומי מופעל. שולח דוח...");
  sendDailyReport();
}
