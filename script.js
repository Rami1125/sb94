// URL of your Google Apps Script acting as the API
// ⚠️ חייבים להחליף את זה ב-URL האמיתי של יישום האינטרנט של ה-Google Apps Script שלך עבור פעולות הנתונים הראשיות.
// כדי לקבל את ה-URL הזה:
// 1. עבור לגיליון ה-Google Sheet שלך המקושר ל-Apps Script.
// 2. פתח את עורך ה-Apps Script (תוספים > Apps Script).
// 3. פרוס את הסקריפט כיישום אינטרנט (פריסה > פריסה חדשה > סוג: יישום אינטרנט).
// 4. ודא ש"הפעל כ:" הוא "אני" ו"למי יש גישה:" הוא "כל אחד".
// 5. העתק את ה-URL של יישום האינטרנט והדבק אותו כאן.
const SCRIPT_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxiS3wXwXCyh8xM1EdTiwXy0T-UyBRQgfrnRRis531lTxmgtJIGawfsPeetX5nVJW3V/exec';

// URL של סקריפט Apps Script נפרד לרישום הודעות WhatsApp (⚠️ החלף ב-ID האמיתי של הסקריפט שלך)
// תצטרך פרויקט Apps Script נפרד שפרוס כיישום אינטרנט במיוחד לרישום הודעות WhatsApp.
const WHATSAPP_LOG_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxiS3wXwXCyh8xM1EdTiwXy0T-UyBRQgfrnRRis531lTxmgtJIGawfsPeetX5nVJW3V/exec';

// URL של סקריפט Apps Script נפרד לשליחת מיילים (⚠️ החלף ב-ID האמיתי של הסקריפט שלך)
// תצטרך פרויקט Apps Script נפרד נוסף שפרוס כיישום אינטרנט במיוחד לשליחת מיילים.
// חשוב: זה צריך להיות ה-URL של ה-Apps Script שיצרת עבור הדוחות היומיים!
const EMAIL_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycby5tIOIoIKgL1QrT-8Rx5WpsA_Amu4_vMRnPs6lyD61iBNCggmuXrwcDzqf-pa_TNZ_/exec'; // 🚨🚨🚨 חובה לעדכן!!!

let allOrders = []; // Array containing all loaded orders
let currentEditingOrder = null; // Variable for the order currently being edited
let autoFillData = null; // Data for customer autofill
let charts = {}; // Object to store Chart.js and Leaflet instances
const OVERDUE_THRESHOLD_DAYS = 10; // Days after order date to be considered 'overdue'

// --- Pagination for Main Orders Table ---
const MAIN_TABLE_INITIAL_DISPLAY_LIMIT = 50;
let currentMainTableDisplayCount = MAIN_TABLE_INITIAL_DISPLAY_LIMIT;
let filteredMainOrders = []; // Store currently filtered orders for main table pagination

// --- Modal Utility Functions ---
function openModal(id) {
    document.getElementById(id).classList.add('active');
    if (id === 'order-details-modal' && charts.orderMap) {
        // Invalidate size to ensure Leaflet map renders correctly after modal animation
        charts.orderMap.invalidateSize();
    } else if (id === 'customer-analysis-details-modal') {
         // Logic for timeline animation needs to run here if applicable
    }
}

function closeModal(id) {
    document.getElementById(id).classList.remove('active');
    if (id === 'order-details-modal' && charts.orderMap) {
        // Optionally destroy map instance to free up resources if not needed
        // charts.orderMap.remove(); 
        // delete charts.orderMap;
    }
}

// --- Loader Functions ---
function showLoader() { document.getElementById('loader-overlay').classList.remove('opacity-0', 'pointer-events-none'); }
function hideLoader() { document.getElementById('loader-overlay').classList.add('opacity-0', 'pointer-events-none'); }

// --- Theme Toggle Functions ---
function toggleTheme() {
    const isDarkMode = document.body.classList.toggle('dark');
    localStorage.setItem('theme', isDarkMode ? 'dark' : 'light');
    document.querySelector('#theme-toggle i').className = isDarkMode ? 'fas fa-moon' : 'fas fa-sun';
    drawCharts(); // Redraw dashboard charts to match new theme colors
    if (currentPage === 'reports') {
        filterReports(); // Re-filter and redraw report charts based on current filter
    }
}

function initializeTheme() {
    const savedTheme = localStorage.getItem('theme');
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    if (savedTheme === 'dark' || (!savedTheme && prefersDark)) {
        document.body.classList.add('dark');
        document.querySelector('#theme-toggle i').className = 'fas fa-moon';
    }
    document.getElementById('theme-toggle').addEventListener('click', toggleTheme);
}

// --- Alert Notification Function ---
function showAlert(message, type = 'info') {
    const container = document.getElementById('alert-container');
    const alertItem = document.createElement('div');
    let bgColor, icon, textColor, borderColor;
    switch(type) {
        case 'success': bgColor = 'bg-green-50'; borderColor = 'border-green-500'; textColor = 'text-green-700'; icon = 'fa-check-circle'; break;
        case 'error': bgColor = 'bg-red-50'; borderColor = 'border-red-500'; textColor = 'text-red-700'; icon = 'fa-times-circle'; break;
        case 'warning': bgColor = 'bg-yellow-50'; borderColor = 'border-yellow-500'; textColor = 'text-yellow-700'; icon = 'fa-exclamation-triangle'; break;
        default: bgColor = 'bg-blue-50'; borderColor = 'border-blue-500'; textColor = 'text-blue-700'; icon = 'fa-info-circle'; break;
    }
    alertItem.className = `p-4 rounded-lg border-l-4 shadow-md flex items-center gap-3 transform translate-x-full opacity-0 transition-all duration-500 ease-out ${bgColor} ${borderColor} ${textColor}`;
    alertItem.innerHTML = `<i class="fas ${icon}"></i><p>${message}</p>`;
    container.prepend(alertItem);
    
    setTimeout(() => {
        alertItem.style.transform = 'translateX(0)';
        alertItem.style.opacity = '1';
    }, 100);

    setTimeout(() => {
        alertItem.style.transform = 'translateX(100%)';
        alertItem.style.opacity = '0';
        setTimeout(() => alertItem.remove(), 500);
    }, 5000);
}

// --- API Communication Function (Exponential Backoff) ---
// Added a new optional parameter 'scriptUrl' to allow targeting different Apps Scripts
async function fetchData(action, params = {}, retries = 0, scriptUrl = SCRIPT_WEB_APP_URL) {
    showLoader();
    const urlParams = new URLSearchParams({ action, ...params });
    const url = `${scriptUrl}?${urlParams.toString()}`; // Use provided scriptUrl
    console.log(`[fetchData] Request URL: ${url}`);
    try {
        const response = await fetch(url);
        console.log(`[fetchData] Response status: ${response.status}`);
        const data = await response.json();
        console.log(`[fetchData] Response data:`, data);

        if (!response.ok) {
            const errorMessage = data.message || `שגיאת שרת HTTP: ${response.status}`;
            showAlert(errorMessage, 'error');
            console.error("[fetchData] HTTP error:", errorMessage, data);
            return { success: false, message: errorMessage };
        }

        if (!data.success && data.message && data.message.includes('Service invoked too many times')) {
            const delay = Math.pow(2, retries) * 1000;
            if (retries < 5) {
                console.warn(`[fetchData] Service invoked too many times, retrying in ${delay}ms... (Attempt ${retries + 1})`);
                await new Promise(resolve => setTimeout(resolve, delay));
                return fetchData(action, params, retries + 1, scriptUrl); // Pass scriptUrl on retry
            } else {
                showAlert('השרת עמוס מדי, אנא נסה שוב מאוחר יותר.', 'error');
                return { success: false, message: 'Service too busy' };
            }
        } else if (!data.success) {
            showAlert(data.message || 'פעולה נכשלה בשרת.', 'error');
            console.error("[fetchData] Server-side operation failed:", data.message, data);
            return data;
        }

        return data;
    } catch (error) {
        showAlert('שגיאת תקשורת: לא ניתן להתחבר לשרת.', 'error');
        console.error("[fetchData] Network or parsing error:", error);
        return { success: false, message: error.message };
    } finally {
        hideLoader();
    }
}

// --- WhatsApp Logging and Sending ---
async function logWhatsAppMessage(docId, message) {
    // This URL also needs to be correctly configured in your Google Apps Script for logging
    const url = `${WHATSAPP_LOG_SCRIPT_URL}?action=logMessage&docId=${encodeURIComponent(docId)}&message=${encodeURIComponent(message)}`;
    try {
        const response = await fetch(url);
        const data = await response.json();
        if (data.success) {
            console.log("[logWhatsAppMessage] WhatsApp message logged successfully.");
        } else {
            console.error("[logWhatsAppMessage] Failed to log WhatsApp message:", data.message);
        }
    } catch (error) {
        console.error("[logWhatsAppMessage] Error logging WhatsApp message:", error);
    }
}

function sendWhatsAppMessage() {
    const customerName = document.getElementById('whatsapp-customer-name').value;
    const phoneNumber = document.getElementById('whatsapp-phone-number').value;
    const message = document.getElementById('whatsapp-message-input').value;
    const orderDocId = document.getElementById('details-order-id').textContent; // Use this for logging

    if (!phoneNumber) {
        showAlert('אנא בחר לקוח עם מספר טלפון או הזן מספר.', 'warning');
        return;
    }
    if (!message.trim()) {
        showAlert('ההודעה ריקה. אנא כתוב הודעה או בחר תבנית.', 'warning');
        return;
    }

    const whatsappUrl = `https://wa.me/${phoneNumber}?text=${encodeURIComponent(message)}`;
    window.open(whatsappUrl, '_blank');
    
    showAlert('פותח WhatsApp לשליחת הודעה...', 'info');
    logWhatsAppMessage(orderDocId, message);
}

function openWhatsAppAlertsForOrder(sheetRow) {
    const order = allOrders.find(o => o.sheetRow === sheetRow);
    if (!order) {
        showAlert('הזמנה לא נמצאה.', 'error');
        return;
    }
    
    showPage('whatsapp-alerts');
    document.getElementById('whatsapp-customer-name').value = order['שם לקוח'] || '';
    document.getElementById('whatsapp-phone-number').value = order['טלפון לקוח'] || '';
    document.getElementById('whatsapp-address').value = order['כתובת'] || '';
    // Set the order ID in the details-order-id element to be used for logging
    document.getElementById('details-order-id').textContent = order['תעודה'] || '';

    // Try to pre-select a relevant template
    if (order._effectiveStatus === 'חורג') {
        document.getElementById('message-template-select').value = whatsappTemplates.findIndex(t => t.name.includes('חורגת')).toString();
    } else if (order._daysPassedCalculated >= (OVERDUE_THRESHOLD_DAYS - 2) && order._effectiveStatus === 'פתוח') { 
        document.getElementById('message-template-select').value = whatsappTemplates.findIndex(t => t.name.includes('לפני חריגה')).toString();
    } else {
        document.getElementById('message-template-select').value = ''; // Default
    }
    loadWhatsAppTemplate();
}

// --- Populate Agent Filter ---
function populateAgentFilter() {
    const agentSelect = document.getElementById('filter-agent-select');
    // Save the current selected value
    const currentSelectedAgent = agentSelect.value; 
    
    agentSelect.innerHTML = '<option value="all">כל הסוכנים</option>';
    const uniqueAgents = [...new Set(allOrders.map(order => order['שם סוכן']).filter(Boolean))];
    uniqueAgents.sort().forEach(agent => {
        const option = document.createElement('option');
        option.value = agent;
        option.textContent = agent;
        agentSelect.appendChild(option);
    });

    // Restore the previously selected value, if it still exists
    if ([...uniqueAgents, 'all'].includes(currentSelectedAgent)) {
        agentSelect.value = currentSelectedAgent;
    } else {
        agentSelect.value = 'all'; // Default to 'all' if the previous selection is no longer valid
    }
}

// --- Main Data Loading and Processing ---
async function loadOrders() {
    const response = await fetchData('list', { status: 'all' });
    if (response.success) {
        allOrders = response.data.map(order => {
            const orderDate = new Date(order['תאריך הזמנה']);
            const today = new Date();
            const daysPassed = Math.floor((today - orderDate) / (1000 * 60 * 60 * 24));
            order._daysPassedCalculated = daysPassed;

            if (order['סטטוס'] === 'סגור') {
                order._effectiveStatus = 'סגור';
            } else if (daysPassed >= OVERDUE_THRESHOLD_DAYS) {
                order._effectiveStatus = 'חורג';
            } else {
                order._effectiveStatus = 'פתוח';
            }
            order.sheetRow = parseInt(order.sheetRow);
            order['Kanban Status'] = order['Kanban Status'] || null; 
            return order;
        });
        updateDashboard();
        filterTable(); // This will also handle the initial lazy loading for the main table
        updateContainerInventory();
        renderTreatmentBoard();
        populateAgentFilter(); 
        if (currentPage === 'reports') {
            filterReports(); // Re-filter and display reports if on reports page
        }
        if (currentPage === 'whatsapp-alerts') {
            renderAlertsTable(); // Update alerts if on alerts page
        }
        if (currentPage === 'customer-analysis') {
            populateCustomerAnalysisTable(); // Update customer analysis if on that page
        }
    } else {
        showAlert(response.message || 'שגיאה בטעינת הזמנות.', 'error');
    }
}

/**
 * Refreshes all data in the application by reloading orders.
 * This function is now explicitly defined and globally accessible.
 */
async function refreshData() {
    showAlert('מרענן נתונים...', 'info');
    await loadOrders();
    showAlert('הנתונים רועננו בהצלחה!', 'success');
}


// --- Dashboard Updates ---
function updateDashboard() {
    const openOrders = allOrders.filter(o => o._effectiveStatus === 'פתוח');
    const overdueOrders = allOrders.filter(o => o._effectiveStatus === 'חורג');
    
    const containersInUse = new Set();
    allOrders.filter(o => o._effectiveStatus !== 'סגור').forEach(order => {
        String(order['מספר מכולה ירדה'] || '').split(',').map(c => c.trim()).filter(Boolean).forEach(c => containersInUse.add(c));
        // A container is 'in use' if it was dropped and not yet picked up
        String(order['מספר מכולה עלתה'] || '').split(',').map(c => c.trim()).filter(Boolean).forEach(c => containersInUse.delete(c));
    });
    
    const activeCustomers = new Set(allOrders.filter(o => o._effectiveStatus !== 'סגור').map(o => o['שם לקוח']).filter(Boolean));

    document.getElementById('open-orders-count').textContent = openOrders.length;
    document.getElementById('overdue-orders-count').textContent = overdueOrders.length;
    document.getElementById('containers-in-use').textContent = containersInUse.size;
    document.getElementById('active-customers-count').textContent = activeCustomers.size;
    document.getElementById('overdue-customers-badge').textContent = overdueOrders.length;
    
    const actionTypeCounts = allOrders.reduce((acc, order) => {
        const type = order['סוג פעולה'];
        if (type) {
            acc[type] = (acc[type] || 0) + 1;
        }
        return acc;
    }, {});

    document.getElementById('action-type-הורדה-count').textContent = actionTypeCounts['הורדה'] || 0;
    document.getElementById('action-type-החלפה-count').textContent = actionTypeCounts['החלפה'] || 0;
    document.getElementById('action-type-העלאה-count').textContent = actionTypeCounts['העלאה'] || 0;

    drawCharts();
}

// --- Main Order Table Rendering and Filtering (with Lazy Loading Simulation) ---
function renderOrdersTable(ordersToRender) {
    const tableBody = document.querySelector('#orders-table tbody');
    tableBody.innerHTML = '';
    const noOrdersMessage = document.getElementById('no-main-orders');
    const loadMoreContainer = document.getElementById('orders-load-more-container');

    if (ordersToRender.length === 0) {
        noOrdersMessage.classList.remove('hidden');
        loadMoreContainer.classList.add('hidden');
        return;
    } else {
        noOrdersMessage.classList.add('hidden');
    }

    // Only render up to currentMainTableDisplayCount
    const ordersToDisplay = ordersToRender.slice(0, currentMainTableDisplayCount);

    ordersToDisplay.forEach(order => {
        const row = tableBody.insertRow();
        const actionTypeClass = order['סוג פעולה'] ? `action-type-${order['סוג פעולה']}` : '';
        row.className = `border-b border-[var(--color-border)] transition-colors cursor-pointer ${actionTypeClass}`;
        
        if (order._effectiveStatus === 'חורג') {
            row.classList.add('overdue-subtle-highlight'); // Use new class for row background
        }

        row.dataset.sheetRow = order.sheetRow;
        row.onclick = (e) => {
            if (!e.target.closest('.action-icon-btn, .container-badge, .customer-name-link, .tooltip-container')) { 
                showOrderDetailsModal(order.sheetRow);
            }
        };

        const containersTaken = String(order['מספר מכולה ירדה'] || '').split(',').map(c => c.trim()).filter(Boolean);
        const containersBrought = String(order['מספר מכולה עלתה'] || '').split(',').map(c => c.trim()).filter(Boolean);
        const allContainers = new Set([...containersTaken, ...containersBrought]);
        
        const containerHTML = [...allContainers].filter(Boolean).map(c => {
            const insight = getContainerInsight(c, order.sheetRow);
            const tooltipHtml = insight ? `<div class="tooltip-container"><span class="cursor-help">💡</span><div class="tooltip-content">${insight}</div></div>` : '';
            return `<span class="container-badge inline-block bg-[var(--color-secondary)] text-[var(--color-text-base)] text-xs font-semibold px-2.5 py-0.5 rounded-full cursor-pointer hover:bg-[var(--color-primary)] hover:text-white transition-colors" onclick="event.stopPropagation(); showContainerHistory('${c.trim()}')"><i class="fas fa-box"></i> ${c.trim()} ${tooltipHtml}</span>`;
        }).join(' ');

        const daysPassedHtml = order._effectiveStatus === 'חורג' ?
            `<span class="overdue-text-blinking">${order._daysPassedCalculated || ''}</span>` :
            `${order._daysPassedCalculated || ''}`;

        row.innerHTML = `
            <td class="p-3 font-medium" data-label="תאריך">${formatDate(order['תאריך הזמנה'])}</td>
            <td class="p-3 font-medium" data-label="תעודה">${order['תעודה'] || ''}</td>
            <td class="p-3 font-semibold customer-name-link cursor-pointer hover:text-[var(--color-primary)]" onclick="event.stopPropagation(); showCustomerAnalysisDetailsModal('${order['שם לקוח']}')" data-label="לקוח">${order['שם לקוח'] || ''}</td>
            <td class="p-3 font-medium" data-label="כתובת">${order['כתובת'] || ''}</td>
            <td class="p-3" data-label="סוג פעולה">${order['סוג פעולה'] || ''}</td>
            <td class="p-3" data-label="ימים שעברו">${daysPassedHtml}</td>
            <td class="p-3" data-label="מכולות">${containerHTML}</td>
            <td class="p-3" data-label="סטטוס"><span class="status-${(order._effectiveStatus || '').replace(/[/ ]/g, '-').toLowerCase()}">${order._effectiveStatus || ''}</span></td>
            <td class="p-3 whitespace-nowrap" data-label="פעולות">
                <button class="action-icon-btn whatsapp-btn" onclick="event.stopPropagation(); openWhatsAppAlertsForOrder(${order.sheetRow})" title="שלח WhatsApp"><i class="fab fa-whatsapp text-green-500"></i></button>
                <button class="action-icon-btn" onclick="event.stopPropagation(); openOrderModal('edit', ${order.sheetRow})" title="ערוך"><i class="fas fa-edit text-[var(--color-info)]"></i></button>
                <button class="action-icon-btn" onclick="event.stopPropagation(); duplicateOrder(${order.sheetRow})" title="שכפל"><i class="fas fa-copy text-[var(--color-primary)]"></i></button>
                ${order._effectiveStatus !== 'סגור' ? `<button class="action-icon-btn" onclick="event.stopPropagation(); openCloseOrderModal(${order.sheetRow}, '${order['תעודה']}')" title="סגור הזמנה"><i class="fas fa-check-circle text-[var(--color-success)]"></i></button>` : ''}
                <button class="action-icon-btn" onclick="event.stopPropagation(); openDeleteConfirmModal(${order.sheetRow}, '${order['תעודה']}')" title="מחק"><i class="fas fa-trash text-[var(--color-danger)]"></i></button>
            </td>
        `;
    });

    // Show/hide Load More button for main table
    if (currentMainTableDisplayCount < filteredMainOrders.length) {
        loadMoreContainer.classList.remove('hidden');
    } else {
        loadMoreContainer.classList.add('hidden');
    }
}

function filterTable(statusFilterParam = null, actionTypeFilterParam = null, isExplicitButtonFilter = false) {
    let searchText = document.getElementById('search-input').value.toLowerCase().trim();
    const selectedStatusFilter = document.getElementById('filter-status-select').value;
    const selectedActionTypeFilter = document.getElementById('filter-action-type-select').value;
    const selectedAgentFilter = document.getElementById('filter-agent-select').value;
    let showClosed = document.getElementById('show-closed-orders').checked;

    // Reset pagination for new filter
    currentMainTableDisplayCount = MAIN_TABLE_INITIAL_DISPLAY_LIMIT;

    if (isExplicitButtonFilter) {
        if (statusFilterParam === 'חורג') {
            showClosed = false;
            document.getElementById('show-closed-orders').checked = false;
            document.getElementById('search-input').value = '';
            document.getElementById('filter-status-select').value = 'חורג';
            document.getElementById('filter-action-type-select').value = 'all';
            document.getElementById('filter-agent-select').value = 'all';
        } else if (actionTypeFilterParam) {
            document.getElementById('search-input').value = '';
            document.getElementById('filter-status-select').value = 'all';
            document.getElementById('filter-action-type-select').value = actionTypeFilterParam;
            document.getElementById('filter-agent-select').value = 'all';
            showClosed = document.getElementById('show-closed-orders').checked;
        } else if (statusFilterParam === 'פתוח' || actionTypeFilterParam === 'מכולה בשימוש' || actionTypeFilterParam === 'לקוח פעיל') {
            document.getElementById('search-input').value = '';
            document.getElementById('filter-status-select').value = statusFilterParam || 'all';
            document.getElementById('filter-action-type-select').value = 'all';
            document.getElementById('filter-agent-select').value = 'all';
            showClosed = false;
            document.getElementById('show-closed-orders').checked = false;
        }
    }

    filteredMainOrders = allOrders.filter(order => {
        let matchesSearch = true;
        if (searchText) {
            const isNumericSearch = !isNaN(parseFloat(searchText)) && isFinite(searchText);
            if (isNumericSearch) {
                matchesSearch = String(order['תעודה'] || '').includes(searchText) ||
                                String(order['מספר מכולה ירדה'] || '').includes(searchText) ||
                                String(order['מספר מכולה עלתה'] || '').includes(searchText);
            } else {
                matchesSearch = Object.values(order).some(val => 
                    String(val).toLowerCase().includes(searchText)
                );
            }
        }
        
        let matchesStatus = true;
        const currentStatusFilter = isExplicitButtonFilter && statusFilterParam ? statusFilterParam : selectedStatusFilter;
        if (currentStatusFilter !== 'all') {
            matchesStatus = (order._effectiveStatus === currentStatusFilter);
        }

        let matchesActionType = true;
        const currentActionTypeFilter = isExplicitButtonFilter && actionTypeFilterParam ? actionTypeFilterParam : selectedActionTypeFilter;
        if (currentActionTypeFilter !== 'all') {
            if (currentActionTypeFilter === 'מכולה בשימוש') {
                matchesActionType = (String(order['מספר מכולה ירדה'] || '').split(',').map(c => c.trim()).filter(Boolean).length > 0 && order._effectiveStatus !== 'סגור' && order['סוג פעולה'] !== 'העלאה');
            } else if (currentActionTypeFilter === 'לקוח פעיל') {
                 matchesActionType = (order._effectiveStatus !== 'סגור');
            }
            else {
                matchesActionType = (order['סוג פעולה'] === currentActionTypeFilter);
            }
        }

        let matchesAgent = true;
        if (selectedAgentFilter !== 'all') {
            matchesAgent = (order['שם סוכן'] === selectedAgentFilter);
        }

        let matchesShowClosed = true;
        if (!showClosed) {
            matchesShowClosed = (order._effectiveStatus === 'פתוח' || order._effectiveStatus === 'חורג');
        }
        
        return matchesSearch && matchesStatus && matchesActionType && matchesAgent && matchesShowClosed;
    });
    filteredMainOrders.sort((a,b) => b.sheetRow - a.sheetRow); // Ensure consistent order when filtering
    renderOrdersTable(filteredMainOrders); // Render filtered orders (with lazy loading)
}

function loadMoreOrdersData() {
    currentMainTableDisplayCount += MAIN_TABLE_INITIAL_DISPLAY_LIMIT;
    renderOrdersTable(filteredMainOrders);
}

// --- Input Clear Buttons ---
function clearInput(id) {
    document.getElementById(id).value = '';
    if (id.includes('search')) { // Re-filter if it's a search input
        filterTable();
    } else if (id.includes('report')) { // Re-filter for reports
        filterReports();
    } else if (id.includes('customer-analysis-search')) {
        filterCustomerAnalysis();
    }
}

function clearSelect(id) {
    document.getElementById(id).value = 'all';
    filterTable(); // Re-filter if it's a select filter
}

// --- Rest of the existing functions (formatDate, validateContainerUsage, etc.) ---
function formatDate(dateInput) {
    if (!dateInput) return '';
    const date = new Date(dateInput);
    if (isNaN(date.getTime())) return dateInput;
    return date.toLocaleDateString('he-IL', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

/**
 * Checks if a container number is currently in use by an active order.
 * Considers 'הורדה' and 'החלפה' as taking a container, and 'העלאה' as returning it.
 * @param {string} containerNum The container number to check.
 * @param {number|null} currentOrderSheetRow Optional. The sheetRow of the order currently being edited, to exclude it from the check.
 * @returns {boolean} True if the container is available (not in use by an open/overdue order), false otherwise.
 */
function validateContainerUsage(containerNum, currentOrderSheetRow = null) {
    if (!containerNum) return true; // Empty container number is always valid

    const today = new Date();

    // Find the latest status of this container across all orders
    let containerEvents = allOrders
        .flatMap(order => {
            const events = [];
            const orderDate = new Date(order['תאריך הזמנה']);

            // Ignore the order currently being edited/added for availability check
            if (currentOrderSheetRow !== null && order.sheetRow === currentOrderSheetRow) {
                return [];
            }

            // Containers taken (הורדה/החלפה)
            const containersTakenByOrder = String(order['מספר מכולה ירדה'] || '').split(',').map(c => c.trim()).filter(Boolean);
            if (containersTakenByOrder.includes(containerNum) && (order['סוג פעולה'] === 'הורדה' || order['סוג פעולה'] === 'החלפה')) {
                events.push({ type: 'taken', date: orderDate, status: order._effectiveStatus, orderId: order['תעודה'] });
            }

            // Containers brought (העלאה/החלפה)
            const containersBroughtByOrder = String(order['מספר מכולה עלתה'] || '').split(',').map(c => c.trim()).filter(Boolean);
            if (containersBroughtByOrder.includes(containerNum) && (order['סוג פעולה'] === 'העלאה' || order['סוג פעולה'] === 'החלפה')) {
                // Use close date if available, otherwise order date
                const returnDate = order['תאריך סגירה'] ? new Date(order['תאריך סגירה']) : orderDate;
                events.push({ type: 'returned', date: returnDate, status: order._effectiveStatus, orderId: order['תעודה'] });
            }
            return events;
        })
        .sort((a, b) => a.date.getTime() - b.date.getTime()); // Sort by date ascending

    let isCurrentlyInUse = false;

    for (const event of containerEvents) {
        if (event.type === 'taken') {
            isCurrentlyInUse = true;
        } else if (event.type === 'returned') {
            isCurrentlyInUse = false;
        }
    }

    // Final check based on the last known state and if the order is still open/overdue
    // This handles cases where the latest event is 'taken' and the relevant order is still active
    const lastRelevantOrderForContainer = allOrders
        .filter(order => {
            if (currentOrderSheetRow !== null && order.sheetRow === currentOrderSheetRow) {
                return false;
            }
            const containersTakenByOrder = String(order['מספר מכולה ירדה'] || '').split(',').map(c => c.trim()).filter(Boolean);
            return containersTakenByOrder.includes(containerNum);
        })
        .sort((a,b) => new Date(b['תאריך הזמנה']) - new Date(a['תאריך הזמנה']))[0]; // Get the latest order where this container was taken

    if (lastRelevantOrderForContainer && (lastRelevantOrderForContainer._effectiveStatus === 'פתוח' || lastRelevantOrderForContainer._effectiveStatus === 'חורג')) {
         // If the latest order that took this container is still open/overdue, it is in use
        isCurrentlyInUse = true;
    } else {
        // Otherwise, check if any other open/overdue order has taken this container without a corresponding return
        const openTakenOrders = allOrders.filter(order => {
            if (currentOrderSheetRow !== null && order.sheetRow === currentOrderSheetRow) return false;
            return (order._effectiveStatus === 'פתוח' || order._effectiveStatus === 'חורג') &&
                   String(order['מספר מכולה ירדה'] || '').split(',').map(c => c.trim()).filter(Boolean).includes(containerNum) &&
                   !String(order['מספר מכולה עלתה'] || '').split(',').map(c => c.trim()).filter(Boolean).includes(containerNum);
        });
        if (openTakenOrders.length > 0) {
            isCurrentlyInUse = true;
        }
    }
   
    return !isCurrentlyInUse;
}


/**
 * Provides insights for a given container number.
 * @param {string} containerNum The container number.
 * @param {number} currentOrderSheetRow The sheetRow of the order being displayed/edited.
 * @returns {string|null} An insight message or null if no special insight.
 */
function getContainerInsight(containerNum, currentOrderSheetRow) {
    if (!containerNum) return null;

    const containerOrders = allOrders
        .filter(order => {
            const containersTaken = String(order['מספר מכולה ירדה'] || '').split(',').map(c => c.trim()).filter(Boolean);
            const containersBrought = String(order['מספר מכולה עלתה'] || '').split(',').map(c => c.trim()).filter(Boolean);
            return containersTaken.includes(containerNum) || containersBrought.includes(containerNum);
        })
        .sort((a, b) => new Date(a['תאריך הזמנה']) - new Date(b['תאריך הזמנה']));

    if (containerOrders.length === 0) {
        return 'מכולה זו אינה משויכת לאף הזמנה במערכת.';
    }

    let isCurrentlyOut = false;
    let lastCustomer = '';
    let lastDropDate = null;
    let lastPickupDate = null;

    for (const order of containerOrders) {
        const containersTaken = String(order['מספר מכולה ירדה'] || '').split(',').map(c => c.trim()).filter(Boolean);
        const containersBrought = String(order['מספר מכולה עלתה'] || '').split(',').map(c => c.trim()).filter(Boolean);
        
        if (containersTaken.includes(containerNum) && (order['סוג פעולה'] === 'הורדה' || order['סוג פעולה'] === 'החלפה')) {
            isCurrentlyOut = true;
            lastCustomer = order['שם לקוח'];
            lastDropDate = new Date(order['תאריך הזמנה']);
            lastPickupDate = null; // Reset pickup date
        }
        if (containersBrought.includes(containerNum) && (order['סוג פעולה'] === 'העלאה' || order['סוג פעולה'] === 'החלפה')) {
            isCurrentlyOut = false;
            lastPickupDate = order['תאריך סגירה'] ? new Date(order['תאריך סגירה']) : new Date(order['תאריך הזמנה']);
        }
    }

    if (isCurrentlyOut) {
        if (lastCustomer && lastDropDate) {
            const today = new Date();
            const daysOut = Math.floor((today - lastDropDate) / (1000 * 60 * 60 * 24));
            return `מכולה זו אצל ${lastCustomer} כבר ${daysOut} ימים.`;
        }
        return 'מכולה זו בשימוש אצל לקוח כלשהו.';
    } else if (containerOrders.length > 0 && lastPickupDate) {
        // Check if it was returned and has been available for a while
        const today = new Date();
        const daysAvailable = Math.floor((today - lastPickupDate) / (1000 * 60 * 60 * 24));
        if (daysAvailable > 30) { // Arbitrary threshold for "available for long"
            return `מכולה זו פנויה במלאי כבר ${daysAvailable} ימים.`;
        }
        return 'מכולה זו פנויה במלאי.';
    }

    return null;
}


// --- Order Modal (Add/Edit/Duplicate) Functions ---
function openOrderModal(mode, sheetRow = null) {
    const form = document.getElementById('order-form');
    form.reset();
    currentEditingOrder = null;
    autoFillData = null;
    const saveBtn = document.getElementById('save-order-btn');

    if (mode === 'add') {
        document.getElementById('modal-title').textContent = 'הוסף הזמנה חדשה';
        document.getElementById('תאריך הזמנה').valueAsDate = new Date();
        saveBtn.innerHTML = '<i class="fas fa-save"></i> שמור הזמנה';
        form.onsubmit = async e => {
            e.preventDefault();
            await addOrder(saveBtn);
        };
    } else if (mode === 'edit' && sheetRow) {
        document.getElementById('modal-title').textContent = 'ערוך הזמנה';
        currentEditingOrder = allOrders.find(order => order.sheetRow === sheetRow);
        if (currentEditingOrder) {
            Object.keys(currentEditingOrder).forEach(key => {
                const input = form.elements[key];
                if (input) {
                    if (input.type === 'date' && currentEditingOrder[key]) {
                        input.value = new Date(currentEditingOrder[key]).toISOString().split('T')[0];
                    } else {
                        input.value = currentEditingOrder[key];
                    }
                }
            });
            saveBtn.innerHTML = '<i class="fas fa-save"></i> עדכן הזמנה';
            form.onsubmit = async e => {
                e.preventDefault();
                await editOrder(sheetRow, saveBtn);
            };
        }
    } else if (mode === 'duplicate' && sheetRow) {
        document.getElementById('modal-title').textContent = 'שכפל הזמנה';
        const originalOrder = allOrders.find(order => order.sheetRow === sheetRow);
        if (originalOrder) {
            Object.keys(originalOrder).forEach(key => {
                const input = form.elements[key];
                if (input && !['תעודה', 'תאריך סגירה', 'ימים שעברו', 'מספרי מכולות', '_effectiveStatus', '_daysPassedCalculated', 'sheetRow', 'Kanban Status'].includes(key)) {
                     if (input.type === 'date' && originalOrder[key]) {
                        input.value = new Date(originalOrder[key]).toISOString().split('T')[0];
                    } else {
                        input.value = originalOrder[key];
                    }
                }
            });
            document.getElementById('תאריך הזמנה').valueAsDate = new Date();
            document.getElementById('תעודה').value = '';
            saveBtn.innerHTML = '<i class="fas fa-save"></i> שכפל הזמנה';
            form.onsubmit = async e => {
                e.preventDefault();
                await addOrder(saveBtn); // Treat as new order for saving
            };
        }
    }
    handleActionTypeChange();
    openModal('order-modal');
}

function handleActionTypeChange() {
    const actionType = document.getElementById('סוג פעולה').value;
    const takenDiv = document.getElementById('container-taken-div');
    const broughtDiv = document.getElementById('container-brought-div');
    takenDiv.classList.toggle('hidden', !['הורדה', 'החלפה'].includes(actionType));
    broughtDiv.classList.toggle('hidden', !['העלאה', 'החלפה'].includes(actionType));
}

async function addOrder(btn) {
    btn.innerHTML = '<i class="fas fa-circle-notch fa-spin"></i> שומר...';
    btn.disabled = true;

    const form = document.getElementById('order-form');
    const formData = new FormData(form);
    const orderData = Object.fromEntries(formData.entries());
    
    const requiredFields = ['תאריך הזמנה', 'תעודה', 'שם סוכן', 'שם לקוח', 'כתובת', 'סוג פעולה'];
    for (const field of requiredFields) {
        if (!orderData[field] || String(orderData[field]).trim() === '') {
            showAlert(`שדה חובה חסר: ${field}`, 'error');
            btn.innerHTML = '<i class="fas fa-save"></i> שמור הזמנה';
            btn.disabled = false;
            return;
        }
    }

    if (['הורדה', 'החלפה'].includes(orderData['סוג פעולה'])) {
        const containerTaken = String(orderData['מספר מכולה ירדה'] || '').trim();
        if (containerTaken && !validateContainerUsage(containerTaken)) {
            showAlert(`שימו לב: מכולה ${containerTaken} נראית כבר בשימוש בהזמנה פתוחה אחרת. ודאו שזו הפעולה הרצויה.`, 'warning');
        }
    }

    orderData['סטטוס'] = 'פתוח';
    orderData['Kanban Status'] = null;

    const response = await fetchData('add', { data: JSON.stringify(orderData) });
    if (response.success) {
        showAlert(response.message, 'success');
        closeModal('order-modal');
        await loadOrders();
        
        if (['העלאה', 'החלפה'].includes(orderData['סוג פעולה'])) {
            const containerBrought = String(orderData['מספר מכולה עלתה'] || '').trim();
            if (containerBrought) {
                await closePreviousContainerOrders(containerBrought, orderData['תאריך הזמנה']);
            }
        }
    } else {
        showAlert(response.message || 'שגיאה בהוספת הזמנה', 'error');
    }
    btn.innerHTML = '<i class="fas fa-save"></i> שמור הזמנה';
    btn.disabled = false;
}

async function editOrder(sheetRow, btn) {
    btn.innerHTML = '<i class="fas fa-circle-notch fa-spin"></i> מעדכן...';
    btn.disabled = true;

    const form = document.getElementById('order-form');
    const formData = new FormData(form);
    const updateData = Object.fromEntries(formData.entries());
    
    const requiredFields = ['תאריך הזמנה', 'תעודה', 'שם סוכן', 'שם לקוח', 'כתובת', 'סוג פעולה'];
    for (const field of requiredFields) {
        if (!updateData[field] || String(updateData[field]).trim() === '') {
            showAlert(`שדה חובה חסר: ${field}`, 'error');
            btn.innerHTML = '<i class="fas fa-save"></i> עדכן הזמנה';
            btn.disabled = false;
            return;
        }
    }

    if (['הורדה', 'החלפה'].includes(updateData['סוג פעולה'])) {
        const containerTaken = String(updateData['מספר מכולה ירדה'] || '').trim();
        if (containerTaken && !validateContainerUsage(containerTaken, sheetRow)) {
            showAlert(`מכולה ${containerTaken} כבר בשימוש בהזמנה פתוחה אחרת. אנא ודא שהיא פנויה.`, 'error');
            btn.innerHTML = '<i class="fas fa-save"></i> עדכן הזמנה';
            btn.disabled = false;
            return;
        }
    }

    if (updateData.hasOwnProperty('סטטוס')) {
        delete updateData['סטטוס'];
    }
    if (updateData.hasOwnProperty('Kanban Status')) {
        delete updateData['Kanban Status'];
    }
    
    const response = await fetchData('edit', { id: sheetRow, data: JSON.stringify(updateData) });
    if (response.success) {
        showAlert(response.message, 'success');
        closeModal('order-modal');
        await loadOrders();

        if (['העלאה', 'החלפה'].includes(updateData['סוג פעולה'])) {
            const containerBrought = String(updateData['מספר מכולה עלתה'] || '').trim();
            if (containerBrought) {
                await closePreviousContainerOrders(containerBrought, updateData['תאריך הזמנה']);
            }
        }
    } else {
        showAlert(response.message || 'שגיאה בעדכון הזמנה', 'error');
    }
    btn.innerHTML = '<i class="fas fa-save"></i> עדכן הזמנה';
    btn.disabled = false;
}

async function closePreviousContainerOrders(containerNumber, closeDate) {
    const response = await fetchData('closePreviousContainerOrders', { containerNumber, closeDate });
    if (!response.success) {
        console.error("[closePreviousContainerOrders] Failed to update previous orders for container:", containerNumber, response.message);
    }
}

async function duplicateOrder(sheetRow) {
    openOrderModal('duplicate', sheetRow);
}

function openDeleteConfirmModal(sheetRow, orderId) {
    document.getElementById('delete-order-id').textContent = orderId;
    document.getElementById('confirm-delete-btn').onclick = async () => {
        const btn = document.getElementById('confirm-delete-btn');
        btn.innerHTML = '<i class="fas fa-circle-notch fa-spin"></i> מוחק...';
        btn.disabled = true;
        await deleteOrder(sheetRow);
        btn.innerHTML = 'מחק';
        btn.disabled = false;
    };
    openModal('delete-confirm-modal');
}

async function deleteOrder(sheetRow) {
    const response = await fetchData('delete', { id: sheetRow });
    if (response.success) {
        showAlert(response.message, 'success');
        closeModal('delete-confirm-modal');
        loadOrders();
    } else {
        showAlert(response.message || 'שגיאה במחיקת הזמנה', 'error');
    }
}

// --- Autofill Customer Details ---
function checkCustomerExistenceAndAutofill() {
    const customerName = document.getElementById('שם לקוח').value.trim();
    const address = document.getElementById('כתובת').value.trim();
    const phone = document.getElementById('טלפון לקוח').value.trim();

    if (!customerName && !address && !phone) return;

    let latestOrder = allOrders
        .filter(o => {
            const matchesName = customerName ? o['שם לקוח'] === customerName : true;
            const matchesAddress = address ? o['כתובת'] === address : true;
            const matchesPhone = phone ? o['טלפון לקוח'] === phone : true;
            return matchesName && matchesAddress && matchesPhone;
        })
        .sort((a, b) => new Date(b['תאריך הזמנה']) - new Date(a['תאריך הזמנה']))[0];

    if (latestOrder && !currentEditingOrder) {
        autoFillData = latestOrder;
        document.getElementById('autofill-customer-name-display').textContent = `הלקוח ${latestOrder['שם לקוח']} זוהה!`;
        document.getElementById('autofill-message').innerHTML = `הלקוח <b>${latestOrder['שם לקוח']}</b> זוהה מהזמנה קודמת מתאריך <b>${formatDate(latestOrder['תאריך הזמנה'])}</b>. האם ברצונך למלא את פרטיו אוטומטית?`;
        openModal('autofill-confirm-modal');
    }
}

function confirmAutofill(confirm) {
    if (confirm && autoFillData) {
        Object.keys(autoFillData).forEach(key => {
            const input = document.getElementById(key);
            if (input && !['תעודה', 'תאריך הזמנה', 'תאריך סגירה', 'ימים שעברו', 'מספרי מכולות', '_effectiveStatus', '_daysPassedCalculated', 'sheetRow', 'Kanban Status'].includes(key)) {
                 if (input.type === 'date' && autoFillData[key]) {
                    input.value = new Date(autoFillData[key]).toISOString().split('T')[0];
                } else {
                    input.value = autoFillData[key];
                }
            }
        });
        document.getElementById('תאריך הזמנה').valueAsDate = new Date();
        document.getElementById('תעודה').value = '';
        handleActionTypeChange();
    }
    closeModal('autofill-confirm-modal');
    autoFillData = null;
}

// --- Order Details Modal (with Map Integration) ---
function showOrderDetailsModal(sheetRow) {
    const order = allOrders.find(o => o.sheetRow === sheetRow);
    if (!order) {
        showAlert('פרטי הזמנה לא נמצאו.', 'error');
        return;
    }

    document.getElementById('details-order-id').textContent = order['תעודה'] || 'לא ידוע';
    const detailsContent = document.getElementById('order-details-content');
    
    // Clear previous content but keep the map container
    const mapContainer = document.getElementById('mapid');
    detailsContent.innerHTML = '';
    detailsContent.appendChild(mapContainer);

    const orderDetailsHtml = `
        <p><strong>תאריך הזמנה:</strong> ${formatDate(order['תאריך הזמנה'])}</p>
        <p><strong>סטטוס:</strong> <span class="status-${(order._effectiveStatus || '').replace(/[/ ]/g, '-').toLowerCase()}">${order._effectiveStatus || ''}</span></p>
        <p><strong>סוג פעולה:</strong> ${order['סוג פעולה'] || ''}</p>
        <p><strong>תעודה:</strong> ${order['תעודה'] || ''}</p>
        <p><strong>שם סוכן:</strong> ${order['שם סוכן'] || ''}</p>
        <p><strong>שם לקוח:</strong> ${order['שם לקוח'] || ''}</p>
        <p><strong>טלפון לקוח:</strong> ${order['טלפון לקוח'] || ''}</p>
        <p><strong>כתובת:</strong> ${order['כתובת'] || ''}</p>
        ${order['מספר מכולה ירדה'] ? `<p><strong>מספר מכולה ירדה:</strong> ${order['מספר מכולה ירדה']}</p>` : ''}
        ${order['מספר מכולה עלתה'] ? `<p><strong>מספר מכולה עלתה:</strong> ${order['מספר מכולה עלתה']}</p>` : ''}
        ${order['תאריך סיום צפוי'] ? `<p><strong>תאריך סיום צפוי:</strong> ${formatDate(order['תאריך סיום צפוי'])}</p>` : ''}
        ${order['תאריך סגירה'] ? `<p><strong>תאריך סגירה:</strong> ${formatDate(order['תאריך סגירה'])}</p>` : ''}
        ${order['הערות סגירה'] ? `<p><strong>הערות סגירה:</strong> ${order['הערות סגירה']}</p>` : ''}
        <p><strong>הערות:</strong> ${order['הערות'] || 'אין'}</p>
    `;
    detailsContent.insertAdjacentHTML('afterbegin', orderDetailsHtml); // Insert at the beginning

    // Initialize or update Leaflet Map
    if (charts.orderMap) {
        charts.orderMap.remove(); // Destroy existing map instance to prevent duplicates
    }
    charts.orderMap = L.map('mapid').setView([0, 0], 13); // Default view
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
    }).addTo(charts.orderMap);

    // Geocode the address and set map view/marker
    if (order['כתובת']) {
        const geocodeUrl = `https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(order['כתובת'])}`;
        fetch(geocodeUrl)
            .then(res => res.json())
            .then(geoData => {
                if (geoData && geoData.length > 0) {
                    const lat = parseFloat(geoData[0].lat);
                    const lon = parseFloat(geoData[0].lon);
                    charts.orderMap.setView([lat, lon], 13);
                    L.marker([lat, lon]).addTo(charts.orderMap)
                        .bindPopup(`<b>${order['שם לקוח']}</b><br>${order['כתובת']}`)
                        .openPopup();
                } else {
                    showAlert('כתובת לא נמצאה על המ mapa.', 'warning');
                    console.warn('כתובת לא נמצאה ב-OpenStreetMap:', order['כתובת']);
                }
            })
            .catch(error => {
                showAlert('שגיאה בטעינת המ mapa.', 'error');
                console.error('Error geocoding address:', error);
            })
            .finally(() => {
                charts.orderMap.invalidateSize(); // Important for map rendering in modal
            });
    } else {
        showAlert('אין כתובת זמינה להצגה על המ mapa.', 'info');
        charts.orderMap.invalidateSize(); // Important for map rendering in modal
    }
    openModal('order-details-modal');
}

function editOrderFromDetails() {
    const orderId = document.getElementById('details-order-id').textContent;
    const order = allOrders.find(o => o['תעודה'] === orderId);
    if (order) {
        closeModal('order-details-modal');
        openOrderModal('edit', order.sheetRow);
    } else {
        showAlert('שגיאה: לא ניתן למצוא את ההזמנה לעריכה.', 'error');
    }
}

function deleteOrderFromDetails() {
    const orderId = document.getElementById('details-order-id').textContent;
    const order = allOrders.find(o => o['תעודה'] === orderId);
    if (order) {
        closeModal('order-details-modal');
        openDeleteConfirmModal(order.sheetRow, order['תעודה']);
    } else {
        showAlert('שגיאה: לא ניתן למצוא את ההזמנה למחיקה.', 'error');
    }
}

function duplicateOrderFromDetails() {
    const orderId = document.getElementById('details-order-id').textContent;
    const order = allOrders.find(o => o['תעודה'] === orderId);
    if (order) {
        closeModal('order-details-modal');
        duplicateOrder(order.sheetRow);
    } else {
        showAlert('שגיאה: לא ניתן למצוא את ההזמנה לשכפול.', 'error');
    }
}

function shareOrderDetailsOnWhatsApp() {
    const orderId = document.getElementById('details-order-id').textContent;
    const order = allOrders.find(o => o['תעודה'] === orderId);

    if (!order || !order['טלפון לקוח']) {
        showAlert('אין מספר טלפון זמין ללקוח זה.', 'warning');
        return;
    }

    const message = `
שלום ${order['שם לקוח']},

להלן פרטי הזמנה מספר: *${order['תעודה']}*
תאריך הזמנה: ${formatDate(order['תאריך הזמנה'])}
סוג פעולה: ${order['סוג פעולה']}
סטטוס: *${order._effectiveStatus}*
כתובת: ${order['כתובת']}
${order['מספר מכולה ירדה'] ? `מכולה ירדה: ${order['מספר מכולה ירדה']}\n` : ''}
${order['מספר מכולה עלתה'] ? `מכולה עלתה: ${order['מספר מכולה עלתה']}\n` : ''}
${order['תאריך סיום צפוי'] ? `תאריך סיום צפוי: ${formatDate(order['תאריך סיום צפוי'])}\n` : ''}
${order['הערות'] ? `הערות: ${order['הערות']}\n` : ''}

בברכה,
[שם העסק שלך]
    `.trim();

    const whatsappUrl = `https://wa.me/${order['טלפון לקוח']}?text=${encodeURIComponent(message)}`;
    window.open(whatsappUrl, '_blank');
    showAlert('נפתח WhatsApp לשליחת הודעה.', 'info');
    logWhatsAppMessage(order['תעודה'], message);
}

function printOrderDetails() {
    const orderId = document.getElementById('details-order-id').textContent;
    const order = allOrders.find(o => o['תעודה'] === orderId);

    if (!order) {
        showAlert('פרטי הזמנה לא נמצאו להדפסה.', 'error');
        return;
    }

    let printContent = `
        <div id="print-area" dir="rtl" style="font-family: 'Rubik', sans-serif; padding: 20px; color: #2F4F4F;">
            <h1 style="text-align: center; color: #2E8B57; font-size: 28px; margin-bottom: 30px;">
                כרטיס הזמנה - ${order['תעודה']}
            </h1>
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 30px;">
                <tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">תאריך הזמנה:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${formatDate(order['תאריך הזמנה'])}</td></tr>
                <tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">סטטוס:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;"><span style="color: ${order._effectiveStatus === 'חורג' ? '#D64545' : '#2E8B57'}; font-weight: bold;">${order._effectiveStatus}</span></td></tr>
                <tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">סוג פעולה:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${order['סוג פעולה'] || ''}</td></tr>
                <tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">שם סוכן:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${order['שם סוכן'] || ''}</td></tr>
                <tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">שם לקוח:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${order['שם לקוח'] || ''}</td></tr>
                <tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">טלפון לקוח:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${order['טלפון לקוח'] || ''}</td></tr>
                <tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">כתובת:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${order['כתובת'] || ''}</td></tr>
                ${order['מספר מכולה ירדה'] ? `<tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">מכולה ירדה:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${order['מספר מכולה ירדה']}</td></tr>` : ''}
                ${order['מספר מכולה עלתה'] ? `<tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">מכולה עלתה:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${order['מספר מכולה עלתה']}</td></tr>` : ''}
                ${order['תאריך סיום צפוי'] ? `<tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">תאריך סיום צפוי:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${formatDate(order['תאריך סיום צפוי'])}</td></tr>` : ''}
                ${order['תאריך סגירה'] ? `<tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">תאריך סגירה:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${formatDate(order['תאריך סגירה'])}</td></tr>` : ''}
                ${order['הערות סגירה'] ? `<tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">הערות סגירה:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${order['הערות סגירה']}</td></tr>` : ''}
                <tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">הערות:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${order['הערות'] || 'אין'}</td></tr>
            </table>
            <div style="text-align: center; margin-top: 40px; font-size: 14px; color: #607D8B;">
                <p>דוח זה נוצר בתאריך: ${formatDate(new Date())}</p>
            </div>
        </div>
    `;

    const printWindow = window.open('', '_blank');
    printWindow.document.write(`
        <html>
        <head>
            <title>הדפסת פרטי הזמנה - ${order['תעודה']}</title>
            <link href="https://fonts.googleapis.com/css2?family=Rubik:wght@300;400;500;600;700&display=swap" rel="stylesheet">
        </head>
        <body>
            ${printContent}
        </body>
        </html>
    `);
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
    printWindow.close();
}

// --- Container Inventory Functions ---
function updateContainerInventory() {
    const containersInUseTableBody = document.getElementById('containers-in-use-table').querySelector('tbody');
    const containersAvailableTableBody = document.getElementById('containers-available-table').querySelector('tbody');

    containersInUseTableBody.innerHTML = '';
    containersAvailableTableBody.innerHTML = '';

    const containerStatus = {}; // { containerNum: { inUse: boolean, lastEventDate: Date, currentCustomer: string, currentOrderSheetRow: number } }

    // Process all orders to determine current container status
    allOrders.sort((a,b) => new Date(a['תאריך הזמנה']) - new Date(b['תאריך הזמנה'])).forEach(order => {
        const orderDate = new Date(order['תאריך הזמנה']);
        const effectiveStatus = order._effectiveStatus; // 'פתוח', 'חורג', 'סגור'

        // Containers taken (הורדה/החלפה)
        const containersTaken = String(order['מספר מכולה ירדה'] || '').split(',').map(c => c.trim()).filter(Boolean);
        containersTaken.forEach(c => {
            if (effectiveStatus !== 'סגור') { // If the order taking it is still open/overdue
                containerStatus[c] = { inUse: true, lastEventDate: orderDate, currentCustomer: order['שם לקוח'], currentOrderSheetRow: order.sheetRow };
            } else { // If the order is closed, assume it was returned at some point or replaced
                containerStatus[c] = { inUse: false, lastEventDate: orderDate, currentCustomer: '', currentOrderSheetRow: null };
            }
        });

        // Containers brought (העלאה/החלפה)
        const containersBrought = String(order['מספר מכולה עלתה'] || '').split(',').map(c => c.trim()).filter(Boolean);
        containersBrought.forEach(c => {
            // If a container is brought back, it's available, regardless of the order's overall status
            containerStatus[c] = { inUse: false, lastEventDate: order['תאריך סגירה'] ? new Date(order['תאריך סגירה']) : orderDate, currentCustomer: '', currentOrderSheetRow: null };
        });
    });

    const sortedContainerNumbers = Object.keys(containerStatus).sort();

    sortedContainerNumbers.forEach(containerNum => {
        const status = containerStatus[containerNum];
        if (status.inUse) {
            const row = containersInUseTableBody.insertRow();
            row.className = 'border-b border-[var(--color-border)]';
            row.innerHTML = `
                <td class="p-3 font-medium">${containerNum}</td>
                <td class="p-3">${status.currentCustomer || 'לא ידוע'}</td>
                <td class="p-3">${formatDate(status.lastEventDate)}</td>
                <td class="p-3">
                    <button class="action-icon-btn text-lg" onclick="showContainerHistory('${containerNum}')" title="הצג היסטוריה"><i class="fas fa-history text-[var(--color-info)]"></i></button>
                </td>
            `;
        } else {
            const row = containersAvailableTableBody.insertRow();
            row.className = 'border-b border-[var(--color-border)]';
            row.innerHTML = `
                <td class="p-3 font-medium">${containerNum}</td>
                <td class="p-3">${formatDate(status.lastEventDate)}</td>
            `;
        }
    });
}

function showContainerHistory(containerNumber) {
    const historyTableBody = document.getElementById('container-history-table-body');
    historyTableBody.innerHTML = '';
    document.getElementById('history-container-number').textContent = containerNumber;
    document.getElementById('no-container-history').classList.add('hidden');

    const relevantOrders = allOrders
        .filter(order => {
            const containersTaken = String(order['מספר מכולה ירדה'] || '').split(',').map(c => c.trim()).filter(Boolean);
            const containersBrought = String(order['מספר מכולה עלתה'] || '').split(',').map(c => c.trim()).filter(Boolean);
            return containersTaken.includes(containerNumber) || containersBrought.includes(containerNumber);
        })
        .sort((a, b) => new Date(a['תאריך הזמנה']) - new Date(b['תאריך הזמנה'])); // Sort by order date ascending

    if (relevantOrders.length === 0) {
        document.getElementById('no-container-history').classList.remove('hidden');
    } else {
        relevantOrders.forEach(order => {
            const row = historyTableBody.insertRow();
            const startDate = new Date(order['תאריך הזמנה']);
            const endDate = order['תאריך סגירה'] ? new Date(order['תאריך סגירה']) : (order['תאריך סיום צפוי'] ? new Date(order['תאריך סיום צפוי']) : null);
            
            let durationDays = 'N/A';
            if (startDate && endDate) {
                durationDays = Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24));
            } else if (startDate && order._effectiveStatus !== 'סגור') {
                durationDays = Math.floor((new Date() - startDate) / (1000 * 60 * 60 * 24));
            }

            row.innerHTML = `
                <td class="p-3">${order['תעודה'] || ''}</td>
                <td class="p-3">${order['שם לקוח'] || ''}</td>
                <td class="p-3">${order['כתובת'] || ''}</td>
                <td class="p-3">${order['סוג פעולה'] || ''}</td>
                <td class="p-3">${formatDate(startDate)}</td>
                <td class="p-3">${endDate ? formatDate(endDate) : (order['תאריך סיום צפוי'] ? `${formatDate(order['תאריך סיום צפוי'])} (צפוי)` : 'אין')}</td>
                <td class="p-3">${durationDays}</td>
            `;
        });
    }
    openModal('container-history-modal');
}

// --- Treatment Board (Kanban) Functions ---
function renderTreatmentBoard() {
    const overdueColumn = document.getElementById('column-overdue');
    const inProgressColumn = document.getElementById('column-in-progress');
    const resolvedColumn = document.getElementById('column-resolved');
    const noTreatmentOrdersMessage = document.getElementById('no-treatment-orders');

    // Clear existing items but keep titles
    Array.from(overdueColumn.children).forEach((child, index) => { if (index > 0) child.remove(); });
    Array.from(inProgressColumn.children).forEach((child, index) => { if (index > 0) child.remove(); });
    Array.from(resolvedColumn.children).forEach((child, index) => { if (index > 0) child.remove(); });

    const ordersForBoard = allOrders.filter(order => 
        order._effectiveStatus === 'חורג' || 
        (order._effectiveStatus === 'פתוח' && order['Kanban Status'] === 'in-progress') ||
        (order._effectiveStatus === 'פתוח' && order['Kanban Status'] === 'resolved')
    );

    if (ordersForBoard.length === 0) {
        noTreatmentOrdersMessage.classList.remove('hidden');
        return;
    } else {
        noTreatmentOrdersMessage.classList.add('hidden');
    }

    ordersForBoard.forEach(order => {
        const item = document.createElement('div');
        item.className = `kanban-item card p-4 mb-3 cursor-grab ${order._effectiveStatus === 'חורג' ? 'border-red-500 border-2' : ''}`;
        item.draggable = true;
        item.id = `kanban-order-${order.sheetRow}`;
        item.dataset.sheetRow = order.sheetRow;
        item.ondragstart = drag;

        let statusColor = 'text-[var(--color-primary)]';
        if (order._effectiveStatus === 'חורג') statusColor = 'text-[var(--color-danger)]';
        else if (order['Kanban Status'] === 'in-progress') statusColor = 'text-[var(--color-info)]';
        else if (order['Kanban Status'] === 'resolved') statusColor = 'text-[var(--color-success)]';

        item.innerHTML = `
            <div class="flex items-center justify-between mb-2">
                <span class="font-bold text-lg">${order['תעודה']} - ${order['שם לקוח']}</span>
                <span class="text-sm font-semibold ${statusColor}">${order._effectiveStatus === 'חורג' ? 'חורג' : (order['Kanban Status'] === 'in-progress' ? 'בטיפול' : (order['Kanban Status'] === 'resolved' ? 'טופל' : 'פתוח'))}</span>
            </div>
            <p class="text-sm text-[var(--color-text-muted)]">${order['כתובת']}</p>
            <p class="text-sm text-[var(--color-text-muted)]">פעולה: ${order['סוג פעולה']}</p>
            <p class="text-sm text-[var(--color-text-muted)]">ימים שעברו: <span class="${order._effectiveStatus === 'חורג' ? 'overdue-text-blinking' : ''}">${order._daysPassedCalculated}</span></p>
            <div class="flex justify-end gap-2 mt-3">
                <button class="action-icon-btn" onclick="openWhatsAppAlertsForOrder(${order.sheetRow})" title="שלח WhatsApp"><i class="fab fa-whatsapp text-green-500"></i></button>
                <button class="action-icon-btn" onclick="openOrderModal('edit', ${order.sheetRow})" title="ערוך"><i class="fas fa-edit text-[var(--color-info)]"></i></button>
                <button class="action-icon-btn" onclick="showOrderDetailsModal(${order.sheetRow})" title="פרטים"><i class="fas fa-info-circle text-[var(--color-secondary)]"></i></button>
            </div>
        `;
        if (order._effectiveStatus === 'חורג' || order['Kanban Status'] === 'overdue') {
            overdueColumn.appendChild(item);
        } else if (order['Kanban Status'] === 'in-progress') {
            inProgressColumn.appendChild(item);
        } else if (order['Kanban Status'] === 'resolved') {
            resolvedColumn.appendChild(item);
        } else if (order._effectiveStatus === 'פתוח') { // Default to in-progress if not explicitly set
            inProgressColumn.appendChild(item);
            // Also update the backend for these if they are implicitly moved
            updateKanbanStatus(order.sheetRow, 'in-progress');
        }
    });
}

function allowDrop(ev) {
    ev.preventDefault();
    ev.currentTarget.classList.add('drag-over');
}

function drag(ev) {
    ev.dataTransfer.setData("text", ev.target.dataset.sheetRow);
}

function drop(ev) {
    ev.preventDefault();
    const sheetRow = ev.dataTransfer.getData("text");
    const targetColumnId = ev.currentTarget.id;
    let newKanbanStatus = null;

    if (targetColumnId === 'column-overdue') {
        newKanbanStatus = 'overdue';
    } else if (targetColumnId === 'column-in-progress') {
        newKanbanStatus = 'in-progress';
    } else if (targetColumnId === 'column-resolved') {
        newKanbanStatus = 'resolved';
    }
    
    ev.currentTarget.classList.remove('drag-over');
    updateKanbanStatus(sheetRow, newKanbanStatus);
}

function handleDragEnter(ev) {
    ev.preventDefault();
    ev.currentTarget.classList.add('drag-over');
}

function handleDragLeave(ev) {
    ev.currentTarget.classList.remove('drag-over');
}

async function updateKanbanStatus(sheetRow, newStatus) {
    const order = allOrders.find(o => o.sheetRow == sheetRow);
    if (!order) {
        showAlert('הזמנה לא נמצאה לעדכון סטטוס.', 'error');
        return;
    }

    let actualStatus = order._effectiveStatus; // Keep the core status (Open/Overdue/Closed)

    // Special handling for 'resolved' column drop
    if (newStatus === 'resolved') {
        openCloseOrderModal(sheetRow, order['תעודה'], true); // Open modal to close the order
        return; // Exit, the actual update will happen after modal confirmation
    }

    // Prevent moving a closed order from being 're-opened' implicitly
    if (order._effectiveStatus === 'סגור' && newStatus !== 'resolved') {
        showAlert('לא ניתן להעביר הזמנה סגורה לסטטוס פתוח בלוח זה.', 'warning');
        renderTreatmentBoard(); // Re-render to revert visual change
        return;
    }

    // If an overdue order is moved to 'in-progress', its effective status is still 'חורג'
    // We only update the Kanban Status field
    const updateData = { 'Kanban Status': newStatus };

    const response = await fetchData('edit', { id: sheetRow, data: JSON.stringify(updateData) });
    if (response.success) {
        showAlert(`סטטוס הזמנה ${order['תעודה']} עודכן ל-${newStatus === 'in-progress' ? 'בטיפול' : 'חורג'}.`, 'success');
        await loadOrders(); // Reload and re-render the board to reflect changes
    } else {
        showAlert(response.message || 'שגיאה בעדכון סטטוס קנבן.', 'error');
        renderTreatmentBoard(); // Re-render to revert visual change in case of error
    }
}

function openCloseOrderModal(sheetRow, orderId, fromKanban = false) {
    document.getElementById('close-order-id-display').textContent = orderId;
    document.getElementById('close-order-notes').value = ''; // Clear previous notes
    document.getElementById('confirm-close-order-btn').onclick = async () => {
        const notes = document.getElementById('close-order-notes').value;
        const btn = document.getElementById('confirm-close-order-btn');
        btn.innerHTML = '<i class="fas fa-circle-notch fa-spin"></i> סוגר...';
        btn.disabled = true;
        await closeOrder(sheetRow, notes);
        btn.innerHTML = 'אשר סגירה ✅';
        btn.disabled = false;
    };
    openModal('close-order-modal');
}

async function closeOrder(sheetRow, notes) {
    const orderToClose = allOrders.find(o => o.sheetRow == sheetRow);
    if (!orderToClose) {
        showAlert('הזמנה לא נמצאה לסגירה.', 'error');
        return;
    }

    const updateData = {
        'סטטוס': 'סגור',
        'תאריך סגירה': new Date().toISOString().split('T')[0],
        'הערות סגירה': notes,
        'Kanban Status': 'resolved' // Mark as resolved in Kanban when closed
    };
    
    // If the action type was 'הורדה' and 'מספר מכולה ירדה' exists,
    // this should implicitly mean the container is now 'available'.
    // If it was 'החלפה', the container brought should be made available,
    // and the container taken needs its previous order closed.
    
    // This logic is mostly handled by `closePreviousContainerOrders` called after add/edit,
    // but ensure here that if it's explicitly closed as 'הורדה', the container becomes available.
    // For simplicity in this client-side code, we rely on the Apps Script to handle the container status.

    const response = await fetchData('edit', { id: sheetRow, data: JSON.stringify(updateData) });
    if (response.success) {
        showAlert(response.message, 'success');
        closeModal('close-order-modal');
        await loadOrders(); // Reload all data to update dashboard, tables, and Kanban board
    } else {
        showAlert(response.message || 'שגיאה בסגירת הזמנה.', 'error');
    }
}

// --- WhatsApp Alerts Page Functions ---
const whatsappTemplates = [
    { name: "תזכורת תשלום", template: "שלום [שם לקוח],\nזוהי תזכורת לתשלום עבור הזמנה [תעודה]. אנא טפל בכך בהקדם.\nתודה!" },
    { name: "הזמנה חורגת", template: "שלום [שם לקוח],\nהזמנה מספר [תעודה] בכתובת [כתובת] חורגת מתאריך הסיום הצפוי. אנא צור קשר לתיאום המשך טיפול.\nתודה!" },
    { name: "לפני חריגה", template: "שלום [שם לקוח],\nהזמנה מספר [תעודה] בכתובת [כתובת] מתקרבת לתאריך הסיום הצפוי. נשמח לסייע בתיאום פינוי או הארכה במידת הצורך.\nתודה!" },
    { name: "הודעת סגירה", template: "שלום [שם לקוח],\nהזמנה מספר [תעודה] בכתובת [כתובת] נסגרה בהצלחה. תודה שבחרת בנו!\n[שם סוכן]" },
    { name: "אישור הזמנה חדשה", template: "שלום [שם לקוח],\nהזמנה חדשה מספר [תעודה] עבור [סוג פעולה] בכתובת [כתובת] אושרה. אנו בדרך!\n[שם סוכן]" },
    { name: "בקשת מיקום", template: "שלום [שם לקוח],\nלצורך טיפול בהזמנה [תעודה] נדרש מיקום מדויק. אנא שלח מיקום בווטסאפ.\nתודה!" }
];

function populateWhatsAppTemplates() {
    const select = document.getElementById('message-template-select');
    select.innerHTML = '<option value="">בחר תבנית...</option>';
    whatsappTemplates.forEach((template, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = template.name;
        select.appendChild(option);
    });
}

function loadWhatsAppTemplate() {
    const select = document.getElementById('message-template-select');
    const messageInput = document.getElementById('whatsapp-message-input');
    const selectedIndex = select.value;

    if (selectedIndex === "") {
        messageInput.value = "";
        return;
    }

    const template = whatsappTemplates[parseInt(selectedIndex)];
    if (template) {
        const customerName = document.getElementById('whatsapp-customer-name').value || '[שם לקוח]';
        const orderId = document.getElementById('details-order-id').textContent || '[תעודה]';
        const address = document.getElementById('whatsapp-address').value || '[כתובת]';
        const actionType = currentEditingOrder ? currentEditingOrder['סוג פעולה'] : '[סוג פעולה]';
        const agentName = currentEditingOrder ? currentEditingOrder['שם סוכן'] : '[שם סוכן]';

        let populatedMessage = template.template;
        populatedMessage = populatedMessage.replace(/\[שם לקוח\]/g, customerName);
        populatedMessage = populatedMessage.replace(/\[תעודה\]/g, orderId);
        populatedMessage = populatedMessage.replace(/\[כתובת\]/g, address);
        populatedMessage = populatedMessage.replace(/\[סוג פעולה\]/g, actionType);
        populatedMessage = populatedMessage.replace(/\[שם סוכן\]/g, agentName);

        messageInput.value = populatedMessage;
    }
}

function clearWhatsAppMessage() {
    document.getElementById('whatsapp-message-input').value = '';
    document.getElementById('message-template-select').value = '';
}

function renderAlertsTable() {
    const alertsTableBody = document.getElementById('alerts-table-body');
    alertsTableBody.innerHTML = '';
    document.getElementById('no-alerts-needed').classList.add('hidden');

    const ordersNeedingAlert = allOrders.filter(order => {
        // Include overdue orders
        if (order._effectiveStatus === 'חורג') return true;
        // Include orders nearing overdue (e.g., within 2 days of OVERDUE_THRESHOLD_DAYS)
        if (order._effectiveStatus === 'פתוח' && order._daysPassedCalculated >= (OVERDUE_THRESHOLD_DAYS - 2) && order._daysPassedCalculated < OVERDUE_THRESHOLD_DAYS) return true;
        return false;
    }).sort((a,b) => b._daysPassedCalculated - a._daysPassedCalculated); // Sort by most overdue first

    if (ordersNeedingAlert.length === 0) {
        document.getElementById('no-alerts-needed').classList.remove('hidden');
    } else {
        ordersNeedingAlert.forEach(order => {
            const row = alertsTableBody.insertRow();
            row.className = 'border-b border-[var(--color-border)]';
            row.innerHTML = `
                <td class="p-3 font-medium">${order['תעודה'] || ''}</td>
                <td class="p-3">${order['שם לקוח'] || ''}</td>
                <td class="p-3"><span class="status-${(order._effectiveStatus || '').replace(/[/ ]/g, '-').toLowerCase()}">${order._effectiveStatus || ''}</span></td>
                <td class="p-3">${order._daysPassedCalculated || ''}</td>
                <td class="p-3">
                    <button class="btn btn-primary btn-sm" onclick="openWhatsAppAlertsForOrder(${order.sheetRow})">
                        <i class="fab fa-whatsapp"></i> שלח הודעה
                    </button>
                </td>
            `;
        });
    }
}

// --- Reports Page Functions ---
let reportsChartMonthly = null;
let reportsChartDistribution = null;
let filteredReportOrders = [];
const REPORTS_TABLE_INITIAL_DISPLAY_LIMIT = 20;
let currentReportsTableDisplayCount = REPORTS_TABLE_INITIAL_DISPLAY_LIMIT;

function filterReports() {
    const startDate = document.getElementById('report-start-date').value;
    const endDate = document.getElementById('report-end-date').value;

    filteredReportOrders = allOrders.filter(order => {
        const orderDate = new Date(order['תאריך הזמנה']);
        let matches = true;
        if (startDate) {
            matches = matches && orderDate >= new Date(startDate);
        }
        if (endDate) {
            matches = matches && orderDate <= new Date(endDate);
        }
        return matches;
    });
    
    updateReportSummaries(filteredReportOrders);
    drawReportsCharts(filteredReportOrders);
    renderReportsTable(filteredReportOrders);
}

function updateReportSummaries(orders) {
    const downloads = orders.filter(o => o['סוג פעולה'] === 'הורדה').length;
    const exchanges = orders.filter(o => o['סוג פעולה'] === 'החלפה').length;
    const uploads = orders.filter(o => o['סוג פעולה'] === 'העלאה').length;

    document.getElementById('summary-downloads').textContent = downloads;
    document.getElementById('summary-exchanges').textContent = exchanges;
    document.getElementById('summary-uploads').textContent = uploads;
}

function drawReportsCharts(orders) {
    // Monthly Actions Chart
    const monthlyCounts = orders.reduce((acc, order) => {
        const date = new Date(order['תאריך הזמנה']);
        const monthYear = `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}`;
        if (!acc[monthYear]) {
            acc[monthYear] = { 'הורדה': 0, 'החלפה': 0, 'העלאה': 0 };
        }
        if (order['סוג פעולה']) {
            acc[monthYear][order['סוג פעולה']]++;
        }
        return acc;
    }, {});

    const sortedMonths = Object.keys(monthlyCounts).sort();
    const monthlyDownloads = sortedMonths.map(m => monthlyCounts[m]['הורדה']);
    const monthlyExchanges = sortedMonths.map(m => monthlyCounts[m]['החלפה']);
    const monthlyUploads = sortedMonths.map(m => monthlyCounts[m]['העלאה']);

    if (reportsChartMonthly) reportsChartMonthly.destroy();
    const chartMonthlyCtx = document.getElementById('chart-reports-monthly-actions').getContext('2d');
    reportsChartMonthly = new Chart(chartMonthlyCtx, {
        type: 'bar',
        data: {
            labels: sortedMonths,
            datasets: [
                { label: 'הורדה', data: monthlyDownloads, backgroundColor: 'rgba(76, 175, 80, 0.6)' },
                { label: 'החלפה', data: monthlyExchanges, backgroundColor: 'rgba(255, 193, 7, 0.6)' },
                { label: 'העלאה', data: monthlyUploads, backgroundColor: 'rgba(214, 69, 69, 0.6)' }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { stacked: true },
                y: { stacked: true, beginAtZero: true }
            },
            plugins: { legend: { position: 'bottom' } }
        }
    });

    // Action Distribution Chart
    const distributionCounts = orders.reduce((acc, order) => {
        const type = order['סוג פעולה'];
        if (type) {
            acc[type] = (acc[type] || 0) + 1;
        }
        return acc;
    }, {});

    const distributionLabels = Object.keys(distributionCounts);
    const distributionData = Object.values(distributionCounts);

    if (reportsChartDistribution) reportsChartDistribution.destroy();
    const chartDistributionCtx = document.getElementById('chart-reports-action-distribution').getContext('2d');
    charts.reportsChartDistribution = new Chart(chartDistributionCtx, {
        type: 'doughnut',
        data: {
            labels: distributionLabels,
            datasets: [{
                data: distributionData,
                backgroundColor: [
                    'rgba(76, 175, 80, 0.8)',   // הורדה (Download)
                    'rgba(255, 193, 7, 0.8)',   // החלפה (Exchange)
                    'rgba(214, 69, 69, 0.8)'    // העלאה (Upload)
                ],
                borderColor: [
                    'var(--color-surface)',
                    'var(--color-surface)',
                    'var(--color-surface)'
                ],
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        color: 'var(--color-text-base)' // Text color for legend
                    }
                }
            }
        }
    });
}

function renderReportsTable(ordersToRender) {
    const tableBody = document.querySelector('#reports-orders-table tbody');
    tableBody.innerHTML = '';
    const noOrdersMessage = document.getElementById('no-report-orders');
    const loadMoreContainer = document.getElementById('reports-load-more-container');

    if (ordersToRender.length === 0) {
        noOrdersMessage.classList.remove('hidden');
        loadMoreContainer.classList.add('hidden');
        return;
    } else {
        noOrdersMessage.classList.add('hidden');
    }

    // Only render up to currentReportsTableDisplayCount
    const ordersToDisplay = ordersToRender.slice(0, currentReportsTableDisplayCount);

    ordersToDisplay.forEach(order => {
        const row = tableBody.insertRow();
        row.className = 'border-b border-[var(--color-border)]';
        row.innerHTML = `
            <td class="p-3 font-medium">${formatDate(order['תאריך הזמנה'])}</td>
            <td class="p-3">${order['תעודה'] || ''}</td>
            <td class="p-3">${order['שם לקוח'] || ''}</td>
            <td class="p-3">${order['סוג פעולה'] || ''}</td>
            <td class="p-3">${(order['מספר מכולה ירדה'] || '') + (order['מספר מכולה עלתה'] ? ` / ${order['מספר מכולה עלתה']}` : '')}</td>
            <td class="p-3"><span class="status-${(order._effectiveStatus || '').replace(/[/ ]/g, '-').toLowerCase()}">${order._effectiveStatus || ''}</span></td>
        `;
    });

    // Show/hide Load More button for reports table
    if (currentReportsTableDisplayCount < ordersToRender.length) {
        loadMoreContainer.classList.remove('hidden');
    } else {
        loadMoreContainer.classList.add('hidden');
    }
}

function loadMoreReportOrders() {
    currentReportsTableDisplayCount += REPORTS_TABLE_INITIAL_DISPLAY_LIMIT;
    renderReportsTable(filteredReportOrders);
}

function resetReportFilters() {
    document.getElementById('report-start-date').value = '';
    document.getElementById('report-end-date').value = '';
    currentReportsTableDisplayCount = REPORTS_TABLE_INITIAL_DISPLAY_LIMIT;
    filterReports();
}

/**
 * Manually triggers the sending of the daily report email.
 * This function is now explicitly defined and globally accessible.
 */
async function sendDailyReportEmailManual() {
    showAlert('שולח דוח יומי למייל...', 'info');
    
    // The Apps Script will fetch and process the data itself, so we just send a trigger action.
    // ⚠️ חשוב: ודא שכתובת המייל מוחלפת בכתובת מייל אמיתית לבדיקה,
    // ושה-EMAIL_SCRIPT_URL מוגדר כראוי בראש הקובץ!
    const response = await fetchData(
        'sendDailyReport', 
        { recipientEmail: 'your.actual.email@example.com' }, // 🚨🚨🚨 החלף בכתובת המייל האמיתית שלך לבדיקה!!!
        0, 
        EMAIL_SCRIPT_URL // משתמש ב-URL של סקריפט המייל שהוגדר בראש הקובץ.
    );

    if (response.success) {
        showAlert('דוח יומי נשלח בהצלחה למייל!', 'success');
    } else {
        showAlert(response.message || 'שגיאה בשליחת הדוח היומי למייל.', 'error');
    }
}

/**
 * Placeholder for sending reports by email from the reports page.
 * You might want to implement a more specific report email functionality here.
 */
async function sendReportsByEmail() {
    showAlert('פונקציית שליחת דוחות במייל עדיין בפיתוח...', 'info');
    // Implement logic to gather current report data and send it via Apps Script
    // This could be a more dynamic report based on the current filters in the reports section.
}


// --- Customer Analysis Page Functions ---
let customerAnalysisChart = null; // Chart for customer activity
let currentCustomerAnalysisData = {}; // Stores data for the currently selected customer

function populateCustomerAnalysisTable() {
    const tableBody = document.getElementById('customer-analysis-table-body');
    tableBody.innerHTML = '';
    document.getElementById('no-customer-analysis').classList.add('hidden');

    const customerSummaries = {}; // { customerName: { totalOrders: 0, lastAddress: '', lastPhone: '' } }

    allOrders.forEach(order => {
        const customerName = order['שם לקוח'];
        if (!customerName) return;

        if (!customerSummaries[customerName]) {
            customerSummaries[customerName] = {
                totalOrders: 0,
                lastAddress: '',
                lastPhone: '',
                orders: [] // Store full orders for detailed view
            };
        }
        customerSummaries[customerName].totalOrders++;
        // Always update with the latest address/phone from the current order in the loop
        // assuming the orders are somewhat ordered or any recent one is fine
        customerSummaries[customerName].lastAddress = order['כתובת'] || customerSummaries[customerName].lastAddress;
        customerSummaries[customerName].lastPhone = order['טלפון לקוח'] || customerSummaries[customerName].lastPhone;
        customerSummaries[customerName].orders.push(order);
    });

    const searchText = document.getElementById('customer-analysis-search-input').value.toLowerCase().trim();
    const filteredCustomers = Object.keys(customerSummaries).filter(name => 
        name.toLowerCase().includes(searchText) || 
        customerSummaries[name].lastAddress.toLowerCase().includes(searchText) ||
        customerSummaries[name].lastPhone.toLowerCase().includes(searchText) ||
        customerSummaries[name].orders.some(order => String(order['תעודה']).toLowerCase().includes(searchText))
    ).sort();

    if (filteredCustomers.length === 0) {
        document.getElementById('no-customer-analysis').classList.remove('hidden');
    } else {
        filteredCustomers.forEach(customerName => {
            const summary = customerSummaries[customerName];
            const row = tableBody.insertRow();
            row.className = 'border-b border-[var(--color-border)] cursor-pointer';
            row.onclick = () => showCustomerAnalysisDetailsModal(customerName);
            row.innerHTML = `
                <td class="p-3 font-semibold">${customerName}</td>
                <td class="p-3">${summary.lastAddress}</td>
                <td class="p-3">${summary.lastPhone}</td>
                <td class="p-3 text-center">${summary.totalOrders}</td>
                <td class="p-3 whitespace-nowrap">
                    <button class="action-icon-btn text-lg" onclick="event.stopPropagation(); showCustomerAnalysisDetailsModal('${customerName}')" title="הצג פרטים"><i class="fas fa-info-circle text-[var(--color-info)]"></i></button>
                    <button class="action-icon-btn text-lg" onclick="event.stopPropagation(); openWhatsAppAlertsForCustomer('${customerName}', '${summary.lastPhone}', '${summary.lastAddress}')" title="שלח WhatsApp"><i class="fab fa-whatsapp text-green-500"></i></button>
                    <button class="action-icon-btn text-lg" onclick="event.stopPropagation(); printCustomerSummary('${customerName}')" title="הדפס סיכום"><i class="fas fa-print text-[var(--color-secondary)]"></i></button>
                </td>
            `;
        });
    }
}

function filterCustomerAnalysis() {
    populateCustomerAnalysisTable(); // Re-render table based on search input
}

function openWhatsAppAlertsForCustomer(customerName, phoneNumber, address) {
    showPage('whatsapp-alerts');
    document.getElementById('whatsapp-customer-name').value = customerName || '';
    document.getElementById('whatsapp-phone-number').value = phoneNumber || '';
    document.getElementById('whatsapp-address').value = address || '';
    document.getElementById('message-template-select').value = ''; // Clear template selection
    document.getElementById('whatsapp-message-input').value = ''; // Clear message
    document.getElementById('details-order-id').textContent = ''; // Clear order ID for logging if not specific to order
    loadWhatsAppTemplate(); // Load default empty template
}

function printCustomerSummary(customerName) {
    const customerOrders = allOrders.filter(o => o['שם לקוח'] === customerName).sort((a,b) => new Date(b['תאריך הזמנה']) - new Date(a['תאריך הזמנה']));
    if (customerOrders.length === 0) {
        showAlert('אין נתונים להדפסה עבור לקוח זה.', 'warning');
        return;
    }

    const summary = {};
    let lastAddress = '';
    let lastPhone = '';
    let totalOpenOrders = 0;
    let totalClosedOrders = 0;
    let totalOverdueOrders = 0;

    customerOrders.forEach(order => {
        if (order._effectiveStatus === 'פתוח') totalOpenOrders++;
        else if (order._effectiveStatus === 'סגור') totalClosedOrders++;
        else if (order._effectiveStatus === 'חורג') totalOverdueOrders++;
        lastAddress = order['כתובת'] || lastAddress;
        lastPhone = order['טלפון לקוח'] || lastPhone;
    });

    let ordersHtml = customerOrders.map(order => `
        <tr>
            <td style="padding: 8px; border: 1px solid #ddd; text-align: right;">${formatDate(order['תאריך הזמנה'])}</td>
            <td style="padding: 8px; border: 1px solid #ddd; text-align: right;">${order['תעודה'] || ''}</td>
            <td style="padding: 8px; border: 1px solid #ddd; text-align: right;">${order['סוג פעולה'] || ''}</td>
            <td style="padding: 8px; border: 1px solid #ddd; text-align: right;"><span style="color: ${order._effectiveStatus === 'חורג' ? '#D64545' : (order._effectiveStatus === 'פתוח' ? '#2E8B57' : '#607D8B')};">${order._effectiveStatus || ''}</span></td>
            <td style="padding: 8px; border: 1px solid #ddd; text-align: right;">${order['מספר מכולה ירדה'] || 'N/A'}</td>
            <td style="padding: 8px; border: 1px solid #ddd; text-align: right;">${order['מספר מכולה עלתה'] || 'N/A'}</td>
        </tr>
    `).join('');


    let printContent = `
        <div id="print-area" dir="rtl" style="font-family: 'Rubik', sans-serif; padding: 20px; color: #2F4F4F;">
            <h1 style="text-align: center; color: #2E8B57; font-size: 28px; margin-bottom: 30px;">
                סיכום לקוח - ${customerName}
            </h1>
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 30px;">
                <tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">כתובת אחרונה:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${lastAddress}</td></tr>
                <tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">טלפון:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${lastPhone}</td></tr>
                <tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">סה"כ הזמנות:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${customerOrders.length}</td></tr>
                <tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">הזמנות פתוחות:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${totalOpenOrders}</td></tr>
                <tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">הזמנות חורגות:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${totalOverdueOrders}</td></tr>
                <tr><th style="padding: 10px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">הזמנות סגורות:</th><td style="padding: 10px; border: 1px solid #ddd; text-align: right;">${totalClosedOrders}</td></tr>
            </table>

            <h2 style="text-align: center; color: #2F4F4F; font-size: 24px; margin-top: 40px; margin-bottom: 20px;">
                היסטוריית הזמנות מפורטת
            </h2>
            <table style="width: 100%; border-collapse: collapse;">
                <thead>
                    <tr>
                        <th style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">תאריך</th>
                        <th style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">תעודה</th>
                        <th style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">פעולה</th>
                        <th style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">סטטוס</th>
                        <th style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">מכולה ירדה</th>
                        <th style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2; text-align: right;">מכולה עלתה</th>
                    </tr>
                </thead>
                <tbody>
                    ${ordersHtml}
                </tbody>
            </table>
            <div style="text-align: center; margin-top: 40px; font-size: 14px; color: #607D8B;">
                <p>דוח זה נוצר בתאריך: ${formatDate(new Date())}</p>
            </div>
        </div>
    `;

    const printWindow = window.open('', '_blank');
    printWindow.document.write(`
        <html>
        <head>
            <title>סיכום לקוח - ${customerName}</title>
            <link href="https://fonts.googleapis.com/css2?family=Rubik:wght@300;400;500;600;700&display=swap" rel="stylesheet">
        </head>
        <body>
            ${printContent}
        </body>
        </html>
    `);
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
    printWindow.close();
}


function showCustomerAnalysisDetailsModal(customerName) {
    document.getElementById('analysis-details-customer-name').textContent = customerName;
    document.getElementById('analysis-downloads-table-body').innerHTML = '';
    document.getElementById('analysis-uploads-table-body').innerHTML = '';
    document.getElementById('no-downloads').classList.add('hidden');
    document.getElementById('no-uploads').classList.add('hidden');
    document.querySelector('#customer-analysis-details-modal .timeline-container').innerHTML = '<div class="timeline-line"></div>'; // Clear and re-add line

    const customerOrders = allOrders
        .filter(o => o['שם לקוח'] === customerName)
        .sort((a, b) => new Date(a['תאריך הזמנה']) - new Date(b['תאריך הזמנה'])); // Sort by date for timeline and table

    if (customerOrders.length === 0) {
        document.getElementById('no-downloads').classList.remove('hidden');
        document.getElementById('no-uploads').classList.remove('hidden');
        openModal('customer-analysis-details-modal');
        return;
    }

    const downloadsBody = document.getElementById('analysis-downloads-table-body');
    const uploadsBody = document.getElementById('analysis-uploads-table-body');
    const timelineContainer = document.querySelector('#customer-analysis-details-modal .timeline-container');
    const timelineEvents = [];

    customerOrders.forEach(order => {
        const orderDate = new Date(order['תאריך הזמנה']);
        const daysPassed = order._daysPassedCalculated;
        const statusClass = (order._effectiveStatus || '').replace(/[/ ]/g, '-').toLowerCase();

        // Downloads Table
        if (['הורדה', 'החלפה'].includes(order['סוג פעולה'])) {
            const row = downloadsBody.insertRow();
            row.className = `border-b border-[var(--color-border)] status-${statusClass}`;
            row.innerHTML = `
                <td class="p-2">${formatDate(orderDate)}</td>
                <td class="p-2">${order['תעודה'] || ''}</td>
                <td class="p-2">${order['מספר מכולה ירדה'] || ''}</td>
                <td class="p-2"><span class="status-${statusClass}">${order._effectiveStatus || ''}</span></td>
                <td class="p-2">${daysPassed}</td>
                <td class="p-2"><button class="action-icon-btn text-lg" onclick="event.stopPropagation(); showOrderDetailsModal(${order.sheetRow})" title="פרטי הזמנה"><i class="fas fa-info-circle text-[var(--color-secondary)]"></i></button></td>
            `;
            timelineEvents.push({
                date: orderDate,
                type: 'הורדה',
                label: `הורדה: ${order['תעודה']}`,
                sheetRow: order.sheetRow,
                effectiveStatus: order._effectiveStatus
            });
        }

        // Uploads Table
        if (['העלאה', 'החלפה'].includes(order['סוג פעולה'])) {
            const row = uploadsBody.insertRow();
            row.className = `border-b border-[var(--color-border)] status-${statusClass}`;
            row.innerHTML = `
                <td class="p-2">${formatDate(orderDate)}</td>
                <td class="p-2">${order['תעודה'] || ''}</td>
                <td class="p-2">${order['מספר מכולה עלתה'] || ''}</td>
                <td class="p-2"><span class="status-${statusClass}">${order._effectiveStatus || ''}</span></td>
                <td class="p-2">${daysPassed}</td>
                <td class="p-2"><button class="action-icon-btn text-lg" onclick="event.stopPropagation(); showOrderDetailsModal(${order.sheetRow})" title="פרטי הזמנה"><i class="fas fa-info-circle text-[var(--color-secondary)]"></i></button></td>
            `;
            timelineEvents.push({
                date: orderDate,
                type: 'העלאה',
                label: `העלאה: ${order['תעודה']}`,
                sheetRow: order.sheetRow,
                effectiveStatus: order._effectiveStatus
            });
        }
    });

    // Populate Timeline
    timelineEvents.sort((a,b) => a.date - b.date); // Ensure chronological order for timeline

    // Add events to timeline
    timelineEvents.forEach(event => {
        const eventDiv = document.createElement('div');
        eventDiv.className = 'timeline-event';
        let dotColor = 'var(--color-accent)';
        if (event.effectiveStatus === 'חורג') dotColor = 'var(--color-danger)';
        else if (event.effectiveStatus === 'סגור') dotColor = 'var(--color-text-muted)';
        else if (event.effectiveStatus === 'פתוח') dotColor = 'var(--color-success)';

        eventDiv.innerHTML = `
            <span class="timeline-dot" style="background-color: ${dotColor};" onclick="showOrderDetailsModal(${event.sheetRow})"></span>
            <span class="timeline-text" onclick="showOrderDetailsModal(${event.sheetRow})">${formatDate(event.date)} - ${event.label}</span>
        `;
        timelineContainer.appendChild(eventDiv);
    });

    // Add animated arrow at the bottom if there are events
    if (timelineEvents.length > 0) {
        const arrow = document.createElement('div');
        arrow.className = 'timeline-arrow-animated';
        arrow.innerHTML = '<i class="fas fa-arrow-down"></i>';
        timelineContainer.appendChild(arrow);
    }

    document.getElementById('no-downloads').classList.toggle('hidden', downloadsBody.children.length > 0);
    document.getElementById('no-uploads').classList.toggle('hidden', uploadsBody.children.length > 0);

    openModal('customer-analysis-details-modal');
}

// --- Chart.js Initialization and Drawing ---
function drawCharts() {
    const primaryColor = getComputedStyle(document.documentElement).getPropertyValue('--color-primary');
    const dangerColor = getComputedStyle(document.documentElement).getPropertyValue('--color-danger');
    const successColor = getComputedStyle(document.documentElement).getPropertyValue('--color-success');
    const warningColor = getComputedStyle(document.documentElement).getPropertyValue('--color-warning');
    const secondaryColor = getComputedStyle(document.documentElement).getPropertyValue('--color-secondary');
    const textBaseColor = getComputedStyle(document.documentElement).getPropertyValue('--color-text-base');

    // Chart: Containers in Use by Customer (Bar Chart)
    const containersByCustomer = allOrders.filter(o => o._effectiveStatus !== 'סגור').reduce((acc, order) => {
        const customer = order['שם לקוח'];
        const containersTaken = String(order['מספר מכולה ירדה'] || '').split(',').map(c => c.trim()).filter(Boolean);
        if (customer) {
            acc[customer] = (acc[customer] || 0) + containersTaken.length;
        }
        return acc;
    }, {});

    const customers = Object.keys(containersByCustomer);
    const containerCounts = Object.values(containersByCustomer);

    if (charts.containersByCustomerChart) charts.containersByCustomerChart.destroy();
    const ctxContainersByCustomer = document.getElementById('chart-containers-by-customer').getContext('2d');
    charts.containersByCustomerChart = new Chart(ctxContainersByCustomer, {
        type: 'bar',
        data: {
            labels: customers,
            datasets: [{
                label: 'מספר מכולות בשימוש',
                data: containerCounts,
                backgroundColor: primaryColor,
                borderColor: primaryColor,
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { color: textBaseColor }
                },
                x: {
                    ticks: { color: textBaseColor }
                }
            },
            plugins: {
                legend: {
                    labels: { color: textBaseColor }
                }
            }
        }
    });

    // Chart: Order Status Distribution (Pie Chart)
    const statusCounts = allOrders.reduce((acc, order) => {
        acc[order._effectiveStatus] = (acc[order._effectiveStatus] || 0) + 1;
        return acc;
    }, { 'פתוח': 0, 'חורג': 0, 'סגור': 0 });

    const statusLabels = ['פתוח', 'חורג', 'סגור'];
    const statusData = statusLabels.map(label => statusCounts[label]);
    const statusColors = [successColor, dangerColor, secondaryColor];

    if (charts.statusPieChart) charts.statusPieChart.destroy();
    const ctxStatusPie = document.getElementById('chart-status-pie').getContext('2d');
    charts.statusPieChart = new Chart(ctxStatusPie, {
        type: 'doughnut',
        data: {
            labels: statusLabels,
            datasets: [{
                data: statusData,
                backgroundColor: statusColors,
                borderColor: getComputedStyle(document.documentElement).getPropertyValue('--color-surface'),
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: { color: textBaseColor }
                }
            }
        }
    });

    // Chart: Action Type Distribution (Bar Chart - Vertical)
    const actionTypeCounts = allOrders.reduce((acc, order) => {
        const type = order['סוג פעולה'];
        if (type) {
            acc[type] = (acc[type] || 0) + 1;
        }
        return acc;
    }, {'הורדה': 0, 'החלפה': 0, 'העלאה': 0});

    const actionTypeLabels = ['הורדה', 'החלפה', 'העלאה'];
    const actionTypeData = actionTypeLabels.map(label => actionTypeCounts[label]);
    const actionTypeColors = [successColor, warningColor, dangerColor];

    if (charts.actionTypeChart) charts.actionTypeChart.destroy();
    const ctxActionType = document.getElementById('chart-action-type').getContext('2d');
    charts.actionTypeChart = new Chart(ctxActionType, {
        type: 'bar',
        data: {
            labels: actionTypeLabels,
            datasets: [{
                label: 'מספר הזמנות',
                data: actionTypeData,
                backgroundColor: actionTypeColors,
                borderColor: actionTypeColors,
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { color: textBaseColor }
                },
                x: {
                    ticks: { color: textBaseColor }
                }
            },
            plugins: {
                legend: {
                    labels: { color: textBaseColor }
                }
            }
        }
    });
}

// --- Page Navigation ---
let currentPage = 'dashboard';
function showPage(pageId) {
    document.querySelectorAll('.page-content').forEach(page => {
        page.classList.add('hidden');
    });
    document.getElementById(`${pageId}-page`).classList.remove('hidden');

    document.querySelectorAll('.nav-tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    document.getElementById(`nav-${pageId}`).classList.add('active');
    currentPage = pageId;

    // Re-load/render data specific to the page when navigated to
    if (pageId === 'container-inventory') {
        updateContainerInventory();
    } else if (pageId === 'treatment-board') {
        renderTreatmentBoard();
    } else if (pageId === 'whatsapp-alerts') {
        populateWhatsAppTemplates();
        renderAlertsTable();
    } else if (pageId === 'reports') {
        resetReportFilters(); // Apply default filters and draw reports
    } else if (pageId === 'customer-analysis') {
        populateCustomerAnalysisTable();
    } else if (pageId === 'dashboard') {
        updateDashboard(); // Ensure dashboard KPIs and charts are up-to-date
    }
}

function scrollToOrdersTable() {
    document.getElementById('orders-table').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function resetTableFilters() {
    document.getElementById('search-input').value = '';
    document.getElementById('filter-status-select').value = 'all';
    document.getElementById('filter-action-type-select').value = 'all';
    document.getElementById('filter-agent-select').value = 'all';
    document.getElementById('show-closed-orders').checked = false;
    filterTable();
}

// --- Scroll to Top Button ---
window.onscroll = function() { scrollFunction() };

function scrollFunction() {
    const scrollToTopBtn = document.getElementById("scroll-to-top-btn");
    if (document.body.scrollTop > 200 || document.documentElement.scrollTop > 200) {
        scrollToTopBtn.style.display = "block";
        scrollToTopBtn.style.opacity = "1";
        scrollToTopBtn.style.transform = "translateY(0)";
    } else {
        scrollToTopBtn.style.opacity = "0";
        scrollToTopBtn.style.transform = "translateY(10px)";
        setTimeout(() => { scrollToTopBtn.style.display = "none"; }, 300);
    }
}

function scrollToTop() {
    window.scrollTo({
        top: 0,
        behavior: 'smooth'
    });
}

// Initial load and setup
document.addEventListener('DOMContentLoaded', async () => {
    initializeTheme();
    await loadOrders(); // Load all data initially
    showPage('dashboard'); // Show dashboard on load
});
