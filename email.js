/**
 * קונפיגורציה של הקובץ.
 */
const SPREADSHEET_ID = '1TPwAP0h05IyzvusybJv3zMKSOKpFNUS9jZtEK5pSgps';
const SHEET_NAME = 'מעקב';
const DEFAULT_RECIPIENT_EMAILS = "ramims@saban94.co.il,rami.msarwa1@gmail.com";
const EMAIL_SUBJECT = "דוח יומי - מערכת CRM מכולות";

function sendDailyReport() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet named "${SHEET_NAME}" not found.`);

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) throw new Error("No data available to create a report.");

  const headers = values[0].map(h => String(h).trim());
  const allOrders = values.slice(1).map(row => {
    let order = {};
    headers.forEach((header, i) => {
      order[header] = row[i] !== undefined && row[i] !== null ? String(row[i]).trim() : '';
    });
    return order;
  });

  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const yesterdayString = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "dd/MM/yyyy");

  const normalize = str => (str ? String(str).trim() : '');

  const openOrders = allOrders.filter(o => normalize(o['סטטוס']) === 'פתוח');
  const overdueOrders = allOrders.filter(o => normalize(o['סטטוס']) === 'חורג');
  const newOrders = allOrders.filter(o => o['תאריך הזמנה'] && formatDateSafe(o['תאריך הזמנה']) === yesterdayString);

  // חישוב מדויק של מספר המכולות
  const totalOpenContainers = sumContainers(openOrders, 'מספר מכולות');
  const totalOverdueContainers = sumContainers(overdueOrders, 'מספר מכולות');
  const totalUsedContainers = totalOpenContainers + totalOverdueContainers;

  Logger.log(`Open Orders: ${openOrders.length}, Containers: ${totalOpenContainers}`);
  Logger.log(`Overdue Orders: ${overdueOrders.length}, Containers: ${totalOverdueContainers}`);

  const htmlBody = generateReportHtml(totalUsedContainers, totalOpenContainers, totalOverdueContainers, openOrders, overdueOrders, newOrders);

  MailApp.sendEmail({
    to: DEFAULT_RECIPIENT_EMAILS,
    subject: EMAIL_SUBJECT,
    htmlBody
  });
}

function sumContainers(orders, columnName) {
  return orders.reduce((sum, o) => {
    let raw = o[columnName] || '0';
    raw = String(raw).replace(/[^0-9.-]/g, ''); // מנקה תווים לא מספריים
    const num = parseFloat(raw);
    return sum + (isNaN(num) ? 0 : num);
  }, 0);
}

function formatDateSafe(value) {
  try {
    return Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), "dd/MM/yyyy");
  } catch (e) {
    return '';
  }
}

function createHtmlTable(orders, title, className) {
  if (orders.length === 0) return `<div class="info-box ${className}"><h3>${title}</h3><p style="text-align: center;">אין נתונים זמינים.</p></div>`;

  const headers = ['תאריך הזמנה', 'תעודה', 'שם סוכן', 'שם לקוח', 'כתובת', 'סוג פעולה', 'ימים שעברו', 'מספר מכולות', 'סטטוס', 'תאריך סיום צפוי'];

  let table = `<div class="table-container ${className}">`;
  table += `<h3 style="text-align: center;">${title}</h3>`;
  table += `<table><thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>`;

  orders.forEach(o => {
    table += `<tr>`;
    headers.forEach(h => {
      let v = o[h] || '';
      if (h.includes('תאריך')) v = formatDateSafe(v) || v;
      let style = '';
      if (h === 'סטטוס') {
        const val = v.trim();
        if (val === 'פתוח') style = 'color:green;font-weight:bold;';
        if (val === 'חורג') style = 'color:red;font-weight:bold;';
        if (val === 'סגור') style = 'color:gray;';
      }
      table += `<td style="${style}">${v}</td>`;
    });
    table += `</tr>`;
  });

  table += `</tbody></table></div>`;
  return table;
}

function generateReportHtml(totalUsed, totalOpen, totalOverdue, openOrders, overdueOrders, newOrders) {
  const todayDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  return `
  <html><head><meta charset="UTF-8">
  <style>
    body{font-family:'Heebo',sans-serif;direction:rtl;background:#f4f6f9;padding:20px;}
    .container{max-width:800px;margin:auto;background:#fff;border-radius:12px;box-shadow:0 4px 10px rgba(0,0,0,0.1);padding:30px;}
    h1,h2,h3{text-align:center;color:#004d99;}
    .summary-flex{display:flex;justify-content:space-around;gap:10px;flex-wrap:wrap;}
    .summary-item{background:#f9f9f9;padding:15px;border-radius:8px;min-width:160px;text-align:center;}
    .summary-item .value{font-size:24px;font-weight:bold;}
    table{width:100%;border-collapse:collapse;margin-top:15px;}
    th,td{border:1px solid #ddd;padding:8px;text-align:right;}
    th{background:#e6f2ff;color:#004d99;}
    .overdue-table th{background:#ffcccc;color:#d93025;}
    .new-orders-table th{background:#dff0d8;color:#10a359;}
    .info-box{background:#fff3cd;color:#856404;padding:15px;border-radius:8px;margin:10px 0;text-align:center;}
  </style>
  </head><body>
  <div class="container">
    <h1>דוח יומי למערכת ניהול מכולות</h1>
    <p style="text-align:center;">תאריך: <b>${todayDate}</b></p>

    <h2>סיכום נתונים</h2>
    <div class="summary-flex">
      <div class="summary-item"><p>סה"כ מכולות בשימוש</p><p class="value">${totalUsed}</p></div>
      <div class="summary-item"><p>מכולות פתוחות</p><p class="value" style="color:green;">${totalOpen}</p></div>
      <div class="summary-item"><p>מכולות חורגות</p><p class="value" style="color:red;">${totalOverdue}</p></div>
    </div>

    ${createHtmlTable(newOrders,'הזמנות מהיום הקודם','new-orders-table')}
    ${createHtmlTable(openOrders,'לקוחות פעילים','open-table')}
    ${createHtmlTable(overdueOrders,'לקוחות חורגים','overdue-table')}

    <p style="text-align:center;margin-top:30px;">בברכה,<br>צוות המכולות</p>
  </div>
  </body></html>`;
}
