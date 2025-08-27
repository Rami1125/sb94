const SPREADSHEET_ID = '1TPwAP0h05IyzvusybJv3zMKSOKpFNUS9jZtEK5pSgps'; 
const ORDERS_SHEET_NAME = 'מעקב'; // שם הגיליון עם נתוני ההזמנות (לפי הקלט שלך)
const WHATSAPP_LOG_SHEET_NAME = 'יומן WhatsApp'; // שם הגיליון ללוג WhatsApp


function doGet(e) {
  var action = e.parameter.action;
  var result = {};

  try {
    switch (action) {
      case 'list':
        // ה-action 'list' יכול כעת לקבל גם פרמטר statusfilter אופציונלי
        var statusFilter = e.parameter.status || 'all'; 
        result = listOrders(statusFilter); 
        break;
      case 'add':
        var orderData = JSON.parse(e.parameter.data);
        result = addOrder(orderData);
        break;
      case 'edit':
        var id = e.parameter.id;
        var updateData = JSON.parse(e.parameter.data);
        result = editOrder(id, updateData);
        break;
      case 'delete':
        var idToDelete = e.parameter.id;
        result = deleteOrder(idToDelete);
        break;
      case 'logMessage': // טיפול בבקשות לתיעוד הודעות WhatsApp
        var docId = e.parameter.docId;
        var message = e.parameter.message;
        result = logWhatsAppMessage(docId, message);
        break;
      case 'closePreviousContainerOrders': // טיפול בסגירת הזמנות קודמות של מכולה
        var containerNumber = e.parameter.containerNumber;
        var closeDate = e.parameter.closeDate;
        result = closePreviousContainerOrders(containerNumber, closeDate);
        break;
      case 'getOrdersByDocIds': // **הפונקציה לחיפוש לפי מספרי תעודות**
        var docIds = JSON.parse(e.parameter.docIds);
        result = getOrdersByDocIds(docIds);
        break;
      case 'getOrdersByCustomerDetails': // **הפונקציה לחיפוש לפי לקוח וכתובת**
        var customerName = e.parameter.customerName;
        var address = e.parameter.address;
        result = getOrdersByCustomerDetails(customerName, address);
        break;
      case 'getOrdersByContainerNumbers': // **הפונקציה לחיפוש לפי מספרי מכולות**
        var containerNumbers = JSON.parse(e.parameter.containerNumbers);
        result = getOrdersByContainerNumbers(containerNumbers);
        break;
      case 'duplicate': // פעולת שכפול הזמנה
        var sheetRowToDuplicate = e.parameter.sheetRow;
        result = duplicateOrder(sheetRowToDuplicate);
        break;
      case 'close': // פעולת סגירת הזמנה
        var sheetRowToClose = e.parameter.sheetRow;
        var notesToClose = e.parameter.notes;
        result = closeOrder(sheetRowToClose, notesToClose);
        break;
      default:
        result = { success: false, message: 'פעולה לא חוקית.' };
    }
  } catch (error) {
    result = { success: false, message: error.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// פונקציית עזר להמרת שורה לאובייקט
function rowToObject(headers, row) {
  var obj = {};
  for (var i = 0; i < headers.length; i++) {
    obj[headers[i]] = row[i];
  }
  return obj;
}

// פונקציית עזר להמרת אובייקט לשורה
function objectToRow(headers, obj) {
  var row = [];
  for (var i = 0; i < headers.length; i++) {
    row.push(obj[headers[i]] !== undefined ? obj[headers[i]] : '');
  }
  return row;
}

/**
 * פונקציה זו שולפת את כל ההזמנות מגיליון ה-Google Sheets.
 * היא קוראת את כל הנתונים בגיליון, ממפה אותם לאובייקטים ומחזירה אותם.
 * כוללת חישוב 'ימים שעברו' וקביעת 'effectiveStatus'.
 * @param {string} statusFilter סטטוס סינון (לדוגמה: 'all', 'פתוח', 'חורג', 'סגור').
 */
function listOrders(statusFilter) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
  if (!sheet) {
    throw new Error('גיליון הזמנות בשם ' + ORDERS_SHEET_NAME + ' לא נמצא.');
  }
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) { // אם יש רק כותרות או שהגיליון ריק
      return { success: true, data: [] };
  }
  var headers = data[0];
  var orders = [];

  for (var i = 1; i < data.length; i++) {
    var order = rowToObject(headers, data[i]);
    order.sheetRow = i + 1; // הוסף את מספר השורה המקורי בגיליון

    // לוגיקה לקביעת effectiveStatus (צריך להיות זהה ל-frontend)
    var orderDate = new Date(order['תאריך הזמנה']);
    var today = new Date();
    today.setHours(0,0,0,0); // נרמול לתחילת היום
    orderDate.setHours(0,0,0,0); // נרמול תאריך הזמנה

    var daysPassed = Math.floor((today - orderDate) / (1000 * 60 * 60 * 24));
    order._daysPassedCalculated = daysPassed; // שמירת ימים שעברו על אובייקט ההזמנה
    
    if (order['סטטוס'] === 'סגור') {
        order._effectiveStatus = 'סגור';
    } else if (daysPassed >= 10) { // אם עברו 10 ימים או יותר, ההזמנה חורגת
        order._effectiveStatus = 'חורג';
    } else {
        order._effectiveStatus = 'פתוח'; // סטטוס ברירת מחדל להזמנות שאינן סגורות ואינן חורגות
    }
    
    // סנן לפי statusFilter אם סופק
    if (statusFilter === 'all' || order._effectiveStatus === statusFilter) {
      orders.push(order);
    }
  }
  return { success: true, data: orders };
}

/**
 * מוסיף הזמנה חדשה לגיליון.
 * @param {Object} orderData אובייקט המכיל את נתוני ההזמנה.
 * @returns {Object} אובייקט המציין הצלחה/כישלון והודעה.
 */
// הוספת הזמנה חדשה (עם טיפול אוטומטי מתקדם)
async function addOrder() {
    // ... (קוד קיים) ...

    // וולידציית שימוש במכולה עבור פעולות 'הורדה' או 'החלפה'
    if (['הורדה', 'החלפה'].includes(orderData['סוג פעולה'])) {
        const containerTaken = String(orderData['מספר מכולה ירדה'] || '').trim();
        if (containerTaken) {
            const isCurrentlyInUse = !validateContainerUsage(containerTaken); // האם היא בשימוש כרגע?

            if (isCurrentlyInUse) {
                // אם המכולה בשימוש, נסגור הזמנות קודמות עבורה
                showAlert(`מכולה ${containerTaken} זוהתה כבשימוש. סוגר הזמנות קודמות פתוחות למכולה זו.`, 'info');
                console.log(`[addOrder] Attempting to close previous orders for container ${containerTaken}.`); // DEBUG
                // קרא לפונקציה לסגירת הזמנות קודמות עבור המכולה שנלקחת
                await closePreviousContainerOrdersForTakeAction(containerTaken, orderData['תאריך הזמנה']);
            }
        }
    }

    orderData['סטטוס'] = 'פתוח'; 
    orderData['Kanban Status'] = null; 

    console.log("[addOrder] Attempting to fetch data with:", orderData); 
    const response = await fetchData('add', { data: JSON.stringify(orderData) });
    if (response.success) {
        showAlert(response.message, 'success');
        closeModal('order-modal');
        console.log("[addOrder] Order added successfully, reloading orders."); 
        await loadOrders(); 
        
        // בצע שינוי סטטוס להזמנות קודמות של המכולה שהועלתה (אם רלוונטי)
        // זה כבר קיים בקוד שלך עבור "העלאה" ו"החלפה"
        if (['העלאה', 'החלפה'].includes(orderData['סוג פעולה'])) {
            const containerBrought = String(orderData['מספר מכולה עלתה'] || '').trim();
            if (containerBrought) {
                console.log(`[addOrder] Closing previous orders for container ${containerBrought}.`); 
                await closePreviousContainerOrders(containerBrought, orderData['תאריך הזמנה']);
            }
        }
    } else {
        showAlert(response.message || 'שגיאה בהוספת הזמנה', 'error');
        console.error("[addOrder] Failed to add order:", response.message); 
    }
}

// פונקציה חדשה לסגירת הזמנות "הורדה" קודמות עבור מכולה ספציפית
// תצטרך להוסיף אותה לסקריפט שלך
async function closePreviousContainerOrdersForTakeAction(containerNumber, newOrderDate) {
    // מצא את כל ההזמנות הפתוחות/חורגות שבהן המכולה הזו ירדה וטרם עלתה
    const ordersToClose = allOrders.filter(order => {
        const containersTakenByOrder = String(order['מספר מכולה ירדה'] || '').split(',').map(c => c.trim()).filter(Boolean);
        const containersBroughtByOrder = String(order['מספר מכולה עלתה'] || '').split(',').map(c => c.trim()).filter(Boolean);

        const isTaken = containersTakenByOrder.includes(containerNumber);
        const isBrought = containersBroughtByOrder.includes(containerNumber);

        return (order._effectiveStatus === 'פתוח' || order._effectiveStatus === 'חורג') && isTaken && !isBrought;
    });

    for (const order of ordersToClose) {
        // בצע קריאה ל-Apps Script לסגור כל הזמנה
        const closeNotes = `נסגר אוטומטית עקב הורדה חדשה למכולה ${containerNumber} בתאריך ${formatDate(newOrderDate)}`;
        console.log(`[closePreviousContainerOrdersForTakeAction] Closing order ${order['תעודה']} (sheetRow: ${order.sheetRow}) for container ${containerNumber}.`); // DEBUG
        await fetchData('close', { sheetRow: order.sheetRow, notes: closeNotes });
    }
}

/**
 * עורך הזמנה קיימת בגיליון.
 * @param {number} sheetRow מספר השורה בגיליון לעריכה.
 * @param {Object} updateData אובייקט עם הנתונים לעדכון.
 * @returns {Object} אובייקט המציין הצלחה/כישלון והודעה.
 */
function editOrder(sheetRow, updateData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
  if (!sheet) {
    throw new Error('גיליון הזמנות בשם ' + ORDERS_SHEET_NAME + ' לא נמצא.');
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  // קבל את כל הנתונים של השורה לפני העדכון כדי לא לאבד נתונים בעמודות לא מעודכנות
  var rowToUpdate = sheet.getRange(parseInt(sheetRow), 1, 1, sheet.getLastColumn()).getValues()[0];
  var currentOrder = rowToObject(headers, rowToUpdate);

  // עדכן רק את השדות שסופקו ב-updateData
  for (var key in updateData) {
    if (updateData.hasOwnProperty(key) && headers.indexOf(key) !== -1) {
      // אם התאריך הוא אובייקט Date (כמו ב-JavaScript), המר אותו לפורמט YYYY-MM-DD
      if (updateData[key] instanceof Date) {
        currentOrder[key] = Utilities.formatDate(updateData[key], SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
      } else {
        currentOrder[key] = updateData[key];
      }
    }
  }
  
  var updatedRow = objectToRow(headers, currentOrder);
  sheet.getRange(parseInt(sheetRow), 1, 1, sheet.getLastColumn()).setValues([updatedRow]);
  return { success: true, message: 'הזמנה עודכנה בהצלחה!' };
}

/**
 * מוחק הזמנה מהגיליון.
 * @param {number} sheetRow מספר השורה בגיליון למחיקה.
 * @returns {Object} אובייקט המציין הצלחה/כישלון והודעה.
 */
function deleteOrder(sheetRow) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
  if (!sheet) {
    throw new Error('גיליון הזמנות בשם ' + ORDERS_SHEET_NAME + ' לא נמצא.');
  }
  sheet.deleteRow(parseInt(sheetRow));
  return { success: true, message: 'הזמנה נמחקה בהצלחה.' };
}

/**
 * פונקציה לתיעוד הודעות WhatsApp בגיליון נפרד.
 * @param {string} docId מזהה המסמך/הזמנה.
 * @param {string} message תוכן ההודעה.
 * @returns {Object} אובייקט המציין הצלחה/כישלון והודעה.
 */
function logWhatsAppMessage(docId, message) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WHATSAPP_LOG_SHEET_NAME);
  if (!sheet) {
    return { success: false, message: 'גיליון יומן WhatsApp לא נמצא.' };
  }
  var timestamp = new Date();
  // לוודא שהכותרות תואמות לגיליון 'יומן WhatsApp' (תאריך ושעה, תעודה, הודעה)
  var row = [timestamp, docId, message];
  sheet.appendRow(row);
  return { success: true, message: 'הודעת WhatsApp נרשמה בהצלחה.' };
}

/**
 * סוגר הזמנות קודמות עבור מכולה ספציפית שהוחזרה/הועלתה.
 * מעדכן את סטטוס ההזמנות ל"סגור" ואת תאריך הסגירה.
 * @param {string} containerNumber מספר המכולה שעבורה רוצים לסגור הזמנות קודמות.
 * @param {string} closeDateStr תאריך הסגירה (בפורמט YYYY-MM-DD) עבור ההזמנות שיסגרו.
 * @returns {Object} אובייקט המציין הצלחה/כישלון והודעה.
 */
function closePreviousContainerOrders(containerNumber, closeDateStr) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
  if (!sheet) {
    throw new Error('גיליון הזמנות בשם ' + ORDERS_SHEET_NAME + ' לא נמצא.');
  }
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) { // אם יש רק כותרות או שהגיליון ריק
      return { success: true, message: 'אין הזמנות לעדכן בגיליון.' };
  }
  var headers = data[0];
  var closeDate = new Date(closeDateStr); // המרת תאריך הסגירה לאובייקט Date

  var docIdCol = headers.indexOf('תעודה');
  var containerTakenCol = headers.indexOf('מספר מכולה ירדה');
  var containerBroughtCol = headers.indexOf('מספר מכולה עלתה');
  var statusCol = headers.indexOf('סטטוס');
  var orderDateCol = headers.indexOf('תאריך הזמנה');
  var closeDateCol = headers.indexOf('תאריך סגירה');

  // בדיקת קיום כל העמודות הנדרשות
  if (docIdCol === -1 || containerTakenCol === -1 || containerBroughtCol === -1 || statusCol === -1 || orderDateCol === -1 || closeDateCol === -1) {
    return { success: false, message: 'חסרה אחת או יותר מעמודות חיוניות (תעודה, מספר מכולה ירדה, מספר מכולה עלתה, סטטוס, תאריך הזמנה, תאריך סגירה).' };
  }

  var updatedRowsCount = 0;
  // לולאה על שורות הנתונים (החל מהשורה השנייה, דילוג על הכותרות)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var currentStatus = row[statusCol];
    var containersTakenInOrder = String(row[containerTakenCol] || '').split(',').map(function(c){ return c.trim(); }).filter(Boolean);
    var containersBroughtInOrder = String(row[containerBroughtCol] || '').split(',').map(function(c){ return c.trim(); }).filter(Boolean);
    var orderDate = new Date(row[orderDateCol]);

    // תנאים לעדכון הזמנה קודמת:
    // 1. המכולה הספציפית ירדה בהזמנה זו.
    // 2. הסטטוס הנוכחי של ההזמנה הוא 'פתוח' או 'חורג'.
    // 3. המכולה הספציפית לא הועלתה בהקשר של הזמנה זו.
    // 4. תאריך ההזמנה קודם לתאריך הסגירה החדש (כדי לא לסגור הזמנות עתידיות בטעות).
    if (containersTakenInOrder.includes(containerNumber) && 
        (currentStatus === 'פתוח' || currentStatus === 'חורג') && 
        !containersBroughtInOrder.includes(containerNumber) &&
        orderDate < closeDate) {
      
      row[statusCol] = 'סגור'; // הגדר סטטוס ל"סגור"
      row[closeDateCol] = Utilities.formatDate(closeDate, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd'); // הגדר תאריך סגירה
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]); // עדכן את השורה בגיליון
      updatedRowsCount++;
    }
  }
  return { success: true, message: 'עודכנו ' + updatedRowsCount + ' הזמנות קודמות עבור מכולה ' + containerNumber };
}


/**
 * מחפש הזמנות בגיליון לפי רשימת מספרי תעודות.
 * @param {Array<string>} docIds מערך של מספרי תעודות לחיפוש.
 * @returns {Object} אובייקט המכיל הצלחה ונתונים (מערך של אובייקטי הזמנות).
 */
function getOrdersByDocIds(docIds) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
  if (!sheet) {
    throw new Error('גיליון הזמנות בשם ' + ORDERS_SHEET_NAME + ' לא נמצא.');
  }
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) { // אם יש רק כותרות או שהגיליון ריק
      return { success: true, data: [] };
  }
  var headers = data[0];
  var docIdCol = headers.indexOf('תעודה');

  if (docIdCol === -1) {
    return { success: false, message: 'עמודת "תעודה" לא נמצאה בגיליון ההזמנות.' };
  }

  var foundOrders = [];
  for (var i = 1; i < data.length; i++) {
    var orderDocId = String(data[i][docIdCol]).trim();
    if (docIds.includes(orderDocId)) {
      var order = rowToObject(headers, data[i]);
      order.sheetRow = i + 1; // הוסף מספר שורה
      foundOrders.push(order);
    }
  }
  return { success: true, data: foundOrders };
}

/**
 * מחפש הזמנות בגיליון לפי שם לקוח וכתובת.
 * @param {string} customerName שם הלקוח לחיפוש (אופציונלי).
 * @param {string} address הכתובת לחיפוש (אופציונלי).
 * @returns {Object} אובייקט המכיל הצלחה ונתונים (מערך של אובייקטי הזמנות).
 */
function getOrdersByCustomerDetails(customerName, address) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
  if (!sheet) {
    throw new Error('גיליון הזמנות בשם ' + ORDERS_SHEET_NAME + ' לא נמצא.');
  }
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) { // אם יש רק כותרות או שהגיליון ריק
      return { success: true, data: [] };
  }
  var headers = data[0];
  var customerNameCol = headers.indexOf('שם לקוח');
  var addressCol = headers.indexOf('כתובת');

  if (customerNameCol === -1 || addressCol === -1) {
    return { success: false, message: 'חסרה אחת מעמודות "שם לקוח" או "כתובת" בגיליון ההזמנות.' };
  }

  var foundOrders = [];
  for (var i = 1; i < data.length; i++) {
    var currentCustomerName = String(data[i][customerNameCol] || '').trim();
    var currentAddress = String(data[i][addressCol] || '').trim();

    // התאמה אם שם הלקוח ריק או תואם, וכנ"ל לכתובת
    var nameMatch = !customerName || currentCustomerName === customerName.trim();
    var addressMatch = !address || currentAddress === address.trim();

    if (nameMatch && addressMatch) {
      var order = rowToObject(headers, data[i]);
      order.sheetRow = i + 1; // הוסף מספר שורה
      foundOrders.push(order);
    }
  }
  return { success: true, data: foundOrders };
}

/**
 * מחפש הזמנות בגיליון לפי רשימת מספרי מכולות (גם כאלו שירדו וגם כאלו שעלו).
 * @param {Array<string>} containerNumbers מערך של מספרי מכולות לחיפוש.
 * @returns {Object} אובייקט המכיל הצלחה ונתונים (מערך של אובייקטי הזמנות).
 */
function getOrdersByContainerNumbers(containerNumbers) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
  if (!sheet) {
    throw new Error('גיליון הזמנות בשם ' + ORDERS_SHEET_NAME + ' לא נמצא.');
  }
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) { // אם יש רק כותרות או שהגיליון ריק
      return { success: true, data: [] };
  }
  var headers = data[0];
  var containerTakenCol = headers.indexOf('מספר מכולה ירדה');
  var containerBroughtCol = headers.indexOf('מספר מכולה עלתה');

  if (containerTakenCol === -1 || containerBroughtCol === -1) {
    return { success: false, message: 'חסרה אחת מעמודות "מספר מכולה ירדה" או "מספר מכולה עלתה" בגיליון ההזמנות.' };
  }

  var foundOrders = [];
  var addedSheetRows = new Set(); // כדי למנוע כפילויות אם אותה מכולה מופיעה בשני שדות באותה הזמנה

  for (var i = 1; i < data.length; i++) {
    var orderRow = i + 1; // מספר השורה בגיליון
    if (addedSheetRows.has(orderRow)) {
        continue; // דלג אם כבר הוספנו את ההזמנה הזו
    }

    var containersTakenInOrder = String(data[i][containerTakenCol] || '').split(',').map(function(c){ return c.trim(); }).filter(Boolean);
    var containersBroughtInOrder = String(data[i][containerBroughtCol] || '').split(',').map(function(c){ return c.trim(); }).filter(Boolean);

    var isMatch = false;
    for (var j = 0; j < containerNumbers.length; j++) {
      var searchContainer = containerNumbers[j];
      if (containersTakenInOrder.includes(searchContainer) || containersBroughtInOrder.includes(searchContainer)) {
        isMatch = true;
        break;
      }
    }

    if (isMatch) {
      var order = rowToObject(headers, data[i]);
      order.sheetRow = orderRow; // הוסף מספר שורה
      foundOrders.push(order);
      addedSheetRows.add(orderRow); // הוסף לשורות שכבר טופלו
    }
  }
  return { success: true, data: foundOrders };
}

/**
 * פונקציה לסגירת הזמנה (מעדכנת סטטוס, תאריך סגירה והערות סיום).
 * משמשת לשינוי סטטוס מלוח הקנבן.
 * @param {number} sheetRow מספר השורה בגיליון לסגירה.
 * @param {string} notes הערות סיום.
 * @returns {Object} אובייקט המציין הצלחה/כישלון והודעה.
 */
function closeOrder(sheetRow, notes) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
  if (!sheet) {
    throw new Error('גיליון הזמנות בשם ' + ORDERS_SHEET_NAME + ' לא נמצא.');
  }

  // וודא ש-sheetRow הוא מספר תקין
  var parsedSheetRow = parseInt(sheetRow, 10); // הוספת רדיקס לוודאות
  
  // בדיקה מפורשת של מספר השורה: אם הוא לא מספר, או שהוא קטן מ-2 (שורת כותרות) או גדול ממספר השורות בגיליון
  if (isNaN(parsedSheetRow) || parsedSheetRow < 2 || parsedSheetRow > sheet.getLastRow()) {
      Logger.log('Received sheetRow for closeOrder: ' + sheetRow + ' (parsed: ' + parsedSheetRow + '). Last row in sheet: ' + sheet.getLastRow()); // רשום ללוג לדיבוג
      throw new Error('שגיאה בסגירת הזמנה: מספר שורה לא חוקי. אנא רענן את העמוד ונסה שוב.');
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rowToUpdate = sheet.getRange(parsedSheetRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var currentOrder = rowToObject(headers, rowToUpdate);

  var statusCol = headers.indexOf('סטטוס');
  var closeDateCol = headers.indexOf('תאריך סגירה');
  var closeNotesCol = headers.indexOf('הערות סיום');

  if (statusCol === -1 || closeDateCol === -1 || closeNotesCol === -1) {
    return { success: false, message: 'חסרה אחת או יותר מעמודות חיוניות (סטטוס, תאריך סגירה, הערות סיום).' };
  }

  currentOrder['סטטוס'] = 'סגור';
  currentOrder['תאריך סגירה'] = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  currentOrder['הערות סיום'] = notes || '';

  var updatedRow = objectToRow(headers, currentOrder);
  sheet.getRange(parsedSheetRow, 1, 1, sheet.getLastColumn()).setValues([updatedRow]);
  return { success: true, message: 'הזמנה נסגרה בהצלחה.' };
}

/**
 * משכפל הזמנה קיימת ומוסיף אותה כשורה חדשה.
 * מאפס שדות מסוימים כדי שתהיה הזמנה חדשה (לדוגמה: סטטוס, תאריכים).
 * @param {number} sheetRow מספר השורה בגיליון לשכפול.
 * @returns {Object} אובייקט המציין הצלחה/כישלון והודעה.
 */
function duplicateOrder(sheetRow) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
  if (!sheet) {
    throw new Error('גיליון הזמנות בשם ' + ORDERS_SHEET_NAME + ' לא נמצא.');
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rowToDuplicate = sheet.getRange(parseInt(sheetRow), 1, 1, sheet.getLastColumn()).getValues()[0];
  var duplicatedOrder = rowToObject(headers, rowToDuplicate);

  // אפס שדות ספציפיים עבור הזמנה משוכפלת
  duplicatedOrder['תעודה'] = 'עותק - ' + duplicatedOrder['תעודה']; 
  duplicatedOrder['סטטוס'] = 'פתוח'; 
  duplicatedOrder['תאריך סגירה'] = ''; 
  duplicatedOrder['הערות סיום'] = ''; 
  duplicatedOrder['ימים שעברו'] = ''; // יחושב מחדש
  duplicatedOrder['ימים שעברו (עלתה)'] = ''; // יחושב מחדש
  duplicatedOrder['תאריך הזמנה'] = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd'); // תאריך היום

  var newRow = objectToRow(headers, duplicatedOrder);
  sheet.appendRow(newRow);
  return { success: true, message: 'הזמנה שוכפלה בהצלחה!' };
}
