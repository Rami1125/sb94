/**
 * קונפיגורציה של הקובץ.
 * יש לעדכן את המזהה ואת שם הגיליון בהתאם לקובץ שלכם.
 */
const SPREADSHEET_ID = '1TPwAP0h05IyzvusybJv3zMKSOKpFNUS9jZtEK5pSgps';
const SHEET_NAME = 'מעקב';
const DEFAULT_RECIPIENT_EMAILS = "ramims@saban94.co.il,rami.msarwa1@gmail.com";
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

  // קוראים את כל הנתונים מהגיליון כדי לאפשר חישובים גמישים
  const dataRange = sheet.getDataRange();
  const values = dataRange.getDisplayValues();

  if (values.length <= 1) {
    Logger.log("No data available to create a report.");
    return;
  }

  const headers = values[0];
  const allOrders = values.slice(1).map(row => {
    let order = {};
    headers.forEach((header, index) => {
      // הסרת רווחים מיותרים מכל הערכים כדי למנוע בעיות התאמה
      order[header] = row[index] ? String(row[index]).trim() : '';
    });
    return order;
  });

  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const yesterdayString = Utilities.formatDate(yesterday, "GMT+2", "dd/MM/yyyy");

  // פילוח הנתונים לקבוצות השונות
  const openOrders = allOrders.filter(order => order['סטטוס'] === 'פתוח' && order['לקוח'] && order['לקוח'].trim() !== '');
  const overdueOrders = allOrders.filter(order => order['סטטוס'] === 'חורג' && order['לקוח'] && order['לקוח'].trim() !== '');
  const newOrders = allOrders.filter(order => {
    const orderDate = order['תאריך הזמנה'];
    return orderDate && Utilities.formatDate(new Date(orderDate), "GMT+2", "dd/MM/yyyy") === yesterdayString;
  });

  // חישובים מדויקים על בסיס נתוני המכולות
  const totalOpenContainers = openOrders.reduce((sum, order) => sum + parseInt(order['מספר מכולות'] || 0), 0);
  const totalOverdueContainers = overdueOrders.reduce((sum, order) => sum + parseInt(order['מספר מכולות'] || 0), 0);
  const totalUsedContainers = totalOpenContainers + totalOverdueContainers;

  // יצירת תוכן ה-HTML של הדוח באמצעות הפונקציה המעוצבת
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
      htmlBody: htmlBody // כאן אנו מוסרים את ה-HTML המעוצב
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
        @import url('https://fonts.googleapis.com/css2?family=Heebo:wght@400;700&display=swap');
        body {
          font-family: 'Heebo', sans-serif;
          direction: rtl;
          text-align: right;
          background-color: #f4f6f9;
          margin: 0;
          padding: 20px;
        }
        .container {
          max-width: 800px;
          margin: 0 auto;
          background-color: #ffffff;
          border-radius: 12px;
          box-shadow: 0 8px 16px rgba(0, 0, 0, 0.1);
          padding: 30px;
          border: 1px solid #e0e0e0;
        }
        h1, h2, h3 {
          font-weight: 700;
          color: #004d99;
          border-bottom: 3px solid #004d99;
          padding-bottom: 8px;
          text-align: center;
        }
        .summary-box {
          background: linear-gradient(135deg, #e6f2ff, #cce6ff);
          border: 1px solid #b3d9ff;
          border-radius: 10px;
          padding: 20px;
          margin-bottom: 30px;
          text-align: center;
          font-weight: bold;
          color: #003366;
        }
        .summary-flex {
            display: flex;
            flex-wrap: wrap;
            justify-content: space-around;
            gap: 15px;
        }
        .summary-item {
            background-color: #ffffff;
            border-radius: 8px;
            padding: 15px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.08);
            min-width: 180px;
            text-align: center;
            transition: transform 0.3s ease;
        }
        .summary-item:hover {
            transform: translateY(-5px);
        }
        .summary-item p {
            margin: 0;
            line-height: 1.5;
        }
        .summary-item .value {
            font-size: 28px;
            font-weight: 700;
            margin-top: 5px;
        }
        .value.used { color: #004d99; }
        .value.open { color: #10a359; }
        .value.overdue { color: #d93025; }
        table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            margin-top: 25px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.05);
            border-radius: 8px;
            overflow: hidden;
        }
        th, td {
            border: 1px solid #e0e0e0;
            padding: 12px;
            text-align: right;
            word-wrap: break-word;
        }
        th {
            background-color: #e6f2ff;
            color: #004d99;
            font-weight: 700;
        }
        .open-table th { background-color: #e6f2ff; }
        .overdue-table th { background-color: #ffcccc; color: #d93025;}
        .new-orders-table th { background-color: #dff0d8; color: #10a359; }
        .table-container {
            margin-bottom: 30px;
            background-color: #fafafa;
            border-radius: 10px;
            padding: 20px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        }
        .info-box {
          background-color: #fff3cd;
          color: #856404;
          border: 1px solid #ffeeba;
          border-radius: 8px;
          padding: 15px;
          text-align: center;
          margin-top: 15px;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>דוח יומי למערכת ניהול מכולות</h1>
        <p style="text-align: center; color: #555;">שלום,</p>
        <p style="text-align: center; color: #555;">זהו דוח עדכון יומי מקיף מיום <b>${todayDate}</b>. הדוח כולל נתוני מכולות בשימוש, הזמנות פתוחות וחורגות, וכן סיכום הזמנות חדשות.</p>
        
        <h2>סיכום נתונים</h2>
        <div class="summary-box">
          <div class="summary-flex">
            <div class="summary-item">
              <p>סה"כ מכולות בשימוש:</p>
              <p class="value used"><b>${totalUsedContainers}</b></p>
            </div>
            <div class="summary-item">
              <p>מכולות בהזמנות פתוחות:</p>
              <p class="value open"><b>${totalOpenContainers}</b></p>
            </div>
            <div class="summary-item">
              <p>מכולות בהזמנות חורגות:</p>
              <p class="value overdue"><b>${totalOverdueContainers}</b></p>
            </div>
          </div>
        </div>
    
        ${newOrdersTable}
        ${openTable}
        ${overdueTable}
    
        <p style="text-align: center; color: #555; margin-top: 30px;">בברכה,</p>
        <p style="text-align: center; color: #555; margin-bottom: 0;">צוות המכולות</p>
      </div>
    </body>
    </html>
  `;
}
