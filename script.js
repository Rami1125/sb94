// URL of your Google Apps Script acting as the API
// REPLACE THIS URL with YOUR OWN deployed Google Apps Script Web App URL for the main data operations.
// To get this URL:
// 1. Go to your Google Sheet linked to the Apps Script.
// 2. Open the Apps Script editor (Extensions > Apps Script).
// 3. Deploy the script as a Web App (Deploy > New deployment > Type: Web app).
// 4. Ensure "Execute as:" is "Me" and "Who has access:" is "Anyone".
// 5. Copy the Web app URL and paste it here.
const SCRIPT_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxiS3wXwXCyh8xM1EdTiwXy0T-UyBRQgfrnRRis531lTxmgtJIGawfsPeetX5nVJW3V/exec';

// URL of a separate Apps Script for WhatsApp message logging (REPLACE WITH YOUR ACTUAL SCRIPT ID)
// You'll need a separate Apps Script project deployed as a Web App specifically for logging WhatsApp messages.
const WHATSAPP_LOG_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbx_YOUR_WHATSAPP_SCRIPT_ID_HERE/exec';

// URL of a separate Apps Script for sending emails (REPLACE WITH YOUR ACTUAL SCRIPT ID)
// You'll need another separate Apps Script project deployed as a Web App specifically for sending emails.
const EMAIL_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbx_YOUR_EMAIL_SCRIPT_ID_HERE/exec';

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
    // Invalidate size for Leaflet maps to ensure correct rendering inside modals
    if (id === 'order-details-modal' && charts.orderMap) {
        charts.orderMap.invalidateSize();
    } else if (id === 'customer-analysis-details-modal' && charts.customerMap) {
        charts.customerMap.invalidateSize();
    }
}

function closeModal(id) {
    document.getElementById(id).classList.remove('active');
    // Optionally destroy map instance to free up resources if not needed
    if (id === 'order-details-modal' && charts.orderMap) {
        // charts.orderMap.remove();
        // delete charts.orderMap;
    } else if (id === 'customer-analysis-details-modal' && charts.customerMap) {
        // charts.customerMap.remove();
        // delete charts.customerMap;
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
    // Re-draw any open modals with charts (e.g., customer analysis, container details)
    if (document.getElementById('customer-analysis-details-modal').classList.contains('active')) {
        const customerName = document.getElementById('analysis-details-customer-name').textContent;
        if (customerName) showCustomerAnalysisDetailsModal(customerName);
    }
    if (document.getElementById('container-details-modal').classList.contains('active')) {
        const containerNum = document.getElementById('details-container-number').textContent;
        if (containerNum) showContainerDetailsModal(containerNum);
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
async function fetchData(action, params = {}, retries = 0, customUrl = SCRIPT_WEB_APP_URL) {
    showLoader();
    const urlParams = new URLSearchParams({ action, ...params });
    const url = `${customUrl}?${urlParams.toString()}`;
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
                return fetchData(action, params, retries + 1, customUrl);
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

// --- Dashboard Updates ---
function updateDashboard() {
    const openOrders = allOrders.filter(o => o._effectiveStatus === 'פתוח');
    const overdueOrders = allOrders.filter(o => o._effectiveStatus === 'חורג');
    
    const containersInUse = new Set();
    allOrders.filter(o => o._effectiveStatus !== 'סגור').forEach(order => {
        // Track containers explicitly taken and not yet returned
        const taken = String(order['מספר מכולה ירדה'] || '').split(',').map(c => c.trim()).filter(Boolean);
        const brought = String(order['מספר מכולה עלתה'] || '').split(',').map(c => c.trim()).filter(Boolean);

        taken.forEach(c => containersInUse.add(c));
        brought.forEach(c => containersInUse.delete(c)); // Remove if brought back
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
            // Prevent opening details modal if interaction is with action buttons, badges, or links
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
            return `<span class="container-badge inline-block bg-[var(--color-secondary)] text-[var(--color-text-base)] text-xs font-semibold px-2.5 py-0.5 rounded-full cursor-pointer hover:bg-[var(--color-primary)] hover:text-white transition-colors" onclick="event.stopPropagation(); showContainerDetailsModal('${c.trim()}')"><i class="fas fa-box"></i> ${c.trim()} ${tooltipHtml}</span>`;
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

// --- Helper Functions (Date, Sorting, etc.) ---
function formatDate(dateInput) {
    if (!dateInput) return '';
    const date = new Date(dateInput);
    if (isNaN(date.getTime())) return dateInput;
    return date.toLocaleDateString('he-IL', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

function parseDateForSort(dateStr) {
    if (!dateStr) return new Date(0); // Return a very early date for empty/null
    const parts = dateStr.split('.'); // Assuming DD.MM.YYYY
    if (parts.length === 3) {
        return new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
    }
    return new Date(dateStr); // Fallback for other formats
}

function sortTable(tableBodyId, columnIndex) {
    const tableBody = document.getElementById(tableBodyId);
    const rows = Array.from(tableBody.rows);
    const isNumeric = [5, 6, 3].includes(columnIndex) && tableBodyId === 'orders-table'; // Days passed, Containers (count), Total Orders
    const isDate = [0].includes(columnIndex); // For date columns
    
    let sortDirection = tableBody.dataset.sortDirection || 'asc'; // Default to ascending
    let currentSortedColumn = tableBody.dataset.sortedColumn;

    // Toggle sort direction if same column clicked again
    if (currentSortedColumn == columnIndex) {
        sortDirection = (sortDirection === 'asc') ? 'desc' : 'asc';
    } else {
        sortDirection = 'asc'; // Default to ascending for new column
    }

    rows.sort((a, b) => {
        let valA = a.cells[columnIndex].textContent.trim();
        let valB = b.cells[columnIndex].textContent.trim();

        if (isDate) {
            valA = parseDateForSort(valA);
            valB = parseDateForSort(valB);
        } else if (isNumeric) {
            valA = parseFloat(valA) || 0;
            valB = parseFloat(valB) || 0;
        } else {
            // For text, use localeCompare for Hebrew sorting
            return sortDirection === 'asc' ? valA.localeCompare(valB, 'he', { sensitivity: 'base' }) : valB.localeCompare(valA, 'he', { sensitivity: 'base' });
        }

        if (valA < valB) {
            return sortDirection === 'asc' ? -1 : 1;
        }
        if (valA > valB) {
            return sortDirection === 'asc' ? 1 : -1;
        }
        return 0;
    });

    // Re-append sorted rows
    rows.forEach(row => tableBody.appendChild(row));

    // Update dataset for next sort
    tableBody.dataset.sortDirection = sortDirection;
    tableBody.dataset.sortedColumn = columnIndex;
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
        .sort((a, b) => new Date(a['תאריך הזמנה']) - new Date(b['תאריך הזמנה'])); // Sort by order date ascending

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
            // Allow override but warn
        }
    }

    orderData['סטטוס'] = 'פתוח';
    orderData['Kanban Status'] = null; // New orders don't start in Kanban status

    const response = await fetchData('add', { data: JSON.stringify(orderData) });
    if (response.success) {
        showAlert(response.message, 'success');
        closeModal('order-modal');
        await loadOrders();
        
        // This is crucial for correctly closing previous related orders for a container.
        // For example, if 'Container X' was "dropped" (הורדה) in order A, and now "picked up" (העלאה) in order B,
        // order A should be marked as closed.
        if (['העלאה', 'החלפה'].includes(orderData['סוג פעולה'])) {
            const containersBrought = String(orderData['מספר מכולה עלתה'] || '').split(',').map(c => c.trim()).filter(Boolean);
            for (const container of containersBrought) {
                if (container) await closePreviousContainerOrders(container, orderData['תאריך הזמנה']);
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

    // Do not allow client to directly update 'סטטוס' or 'Kanban Status' via this form,
    // as these are controlled by specific actions (e.g., close order, Kanban drag-drop)
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

        // This is crucial for correctly closing previous related orders for a container.
        // For example, if 'Container X' was "dropped" (הורדה) in order A, and now "picked up" (העלאה) in order B,
        // order A should be marked as closed.
        if (['העלאה', 'החלפה'].includes(updateData['סוג פעולה'])) {
            const containersBrought = String(updateData['מספר מכולה עלתה'] || '').split(',').map(c => c.trim()).filter(Boolean);
            for (const container of containersBrought) {
                if (container) await closePreviousContainerOrders(container, updateData['תאריך הזמנה']);
            }
        }
    } else {
        showAlert(response.message || 'שגיאה בעדכון הזמנה', 'error');
    }
    btn.innerHTML = '<i class="fas fa-save"></i> עדכן הזמנה';
    btn.disabled = false;
}

// Function to close previous open orders for a specific container
async function closePreviousContainerOrders(containerNumber, closeDate) {
    // Find any existing "open" or "overdue" orders where this container was "ירדה" (dropped)
    // and has not yet been "עלתה" (picked up)
    const ordersToClose = allOrders.filter(order => {
        const containersTaken = String(order['מספר מכולה ירדה'] || '').split(',').map(c => c.trim()).filter(Boolean);
        const containersBrought = String(order['מספר מכולה עלתה'] || '').split(',').map(c => c.trim()).filter(Boolean);
        
        // This order took the container, it's not closed, and hasn't been returned by itself
        return containersTaken.includes(containerNumber) &&
               (order._effectiveStatus === 'פתוח' || order._effectiveStatus === 'חורג') &&
               !containersBrought.includes(containerNumber); // Ensure it wasn't already returned by itself
    });

    for (const order of ordersToClose) {
        if (order.sheetRow) {
            console.log(`[closePreviousContainerOrders] Closing order ${order['תעודה']} (row ${order.sheetRow}) for container ${containerNumber}`);
            const updateData = {
                'סטטוס': 'סגור',
                'תאריך סגירה': closeDate, // Use the date of the new 'העלאה' action
                'הערות סגירה': `נסגר אוטומטית עם החזרת מכולה ${containerNumber} בהזמנה חדשה/מעודכנת.`,
                'Kanban Status': 'resolved' // Mark as resolved in Kanban
            };
            const response = await fetchData('edit', { id: order.sheetRow, data: JSON.stringify(updateData) });
            if (!response.success) {
                console.error(`[closePreviousContainerOrders] Failed to close order ${order['תעודה']}:`, response.message);
                showAlert(`שגיאה בסגירת הזמנה קודמת למכולה ${containerNumber}: ${response.message}`, 'error');
            }
        }
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
            return matchesName && matchesAddr
