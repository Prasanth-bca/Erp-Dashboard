// --- Configuration ---
// BATCH WEBHOOK URL REMOVED
const PROCESSOR_WEBHOOK_URL = 'http://localhost:5678/webhook/cc326b01-5483-4f2a-8a65-cc2df29f80fd'; 
const QUEUE_SHEET_ID = '1oZr8l0fq6P2VTO08MtkW259L87k2NW0WN67PJT7Dz7Q'; // Verify this is your Queue Sheet ID

// Existing Sheet IDs for the dashboard
const SALES_SHEET_ID = '1rfzCkhMnR2VaKy0h6FbTF315l8JQJd2mm07XoyYOKXw';
const PURCHASE_SHEET_ID = '12Ax14fWmK6_fb_WDBAUuS0bGeCqOJz_ZyMjHvRD2Sz4';

// Google API Config
const CLIENT_ID = '1047129115460-6bcckg9gbeod9np02j7uu8nilo1me010.apps.googleusercontent.com';
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';

// Tab names for the dashboard data loading
const SALES_TAB_NAMES = [
    'Request for Quotation', 'Sales Order', 'Sales Invoice', 
    'Customer Payment', 'Debit Note', 'Credit Note'
];
const PURCHASE_TAB_NAMES = [
    'Purchase Order', 'Order Confirmation', 'Delivery Note (DC)', 
    'Principals Invoice', 'Payment Voucher'
];


// --- Global State & Initialization ---
let tokenClient, gapiInited = false, gisInited = false;
let authorizeButton, signoutButton, mainContainer, placeholder, viewModal, viewModalTitle, viewModalBody, viewModalCloseBtn;
let allItemsCache = [];
// PENDING EMAILS CACHE REMOVED
let loadingDiv, errorDiv, emailTable, emailTableBody, refreshInboxBtn;

document.addEventListener('DOMContentLoaded', () => {
    // Get all DOM elements
    authorizeButton = document.getElementById('authorize_button');
    signoutButton = document.getElementById('signout_button');
    mainContainer = document.querySelector('.main-container');
    placeholder = document.getElementById('initial-placeholder');
    viewModal = document.getElementById('view-modal');
    viewModalTitle = document.getElementById('view-modal-title');
    viewModalBody = document.getElementById('view-modal-body');
    viewModalCloseBtn = document.getElementById('view-modal-close-btn');
    loadingDiv = document.getElementById('loading');
    errorDiv = document.getElementById('error');
    emailTable = document.getElementById('email-table');
    emailTableBody = document.getElementById('email-table-body');
    refreshInboxBtn = document.getElementById('refresh-inbox-btn');

    // Setup event listeners
    authorizeButton.onclick = handleAuthClick;
    signoutButton.onclick = handleSignoutClick;
    viewModalCloseBtn.onclick = () => viewModal.style.display = 'none';
    refreshInboxBtn.onclick = fetchEmailList;
    
    // BATCH BUTTON EVENT LISTENER REMOVED

    document.querySelector('.main-tabs').addEventListener('click', (e) => {
        if (e.target.classList.contains('tab-link')) {
            document.querySelectorAll('.tab-link').forEach(tab => tab.classList.remove('active'));
            document.querySelectorAll('.content > .tab-content').forEach(content => content.classList.remove('active'));
            e.target.classList.add('active');
            document.getElementById(e.target.dataset.tab).classList.add('active');
        }
    });
    document.querySelector('.cycle-filter').addEventListener('click', (e) => {
        if (e.target.classList.contains('cycle-btn')) {
            document.querySelectorAll('.cycle-btn').forEach(btn => btn.classList.remove('active'));
            e.target.classList.add('active');
            filterAndRenderData(e.target.dataset.cycle);
        }
    });
    document.querySelector('.content').addEventListener('click', (e) => {
        if (e.target.classList.contains('export-btn')) {
            exportTableToCSV(e.target.dataset.target);
        }
    });
});


// --- Google API Handlers ---
function handleGapiLoad() { gapi.load('client', initializeGapiClient); }

function handleGisLoad() {
    tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: CLIENT_ID, scope: SCOPES,
        callback: async (tokenResponse) => {
            if (tokenResponse.error) {
                placeholder.style.display = 'block';
                mainContainer.style.display = 'none';
                return;
            }
            gapi.client.setToken(tokenResponse);
            placeholder.style.display = 'none';
            mainContainer.style.display = 'block';
            signoutButton.style.display = 'block';
            authorizeButton.style.display = 'none';

            await loadAllData();
            await fetchEmailList();
        },
    });
    gisInited = true; checkAndTrySilentSignIn();
}

async function initializeGapiClient() {
    await gapi.client.init({ discoveryDocs: ["https://sheets.googleapis.com/$discovery/rest?version=v4"] });
    gapiInited = true; checkAndTrySilentSignIn();
}
function checkAndTrySilentSignIn() { if (gapiInited && gisInited) tokenClient.requestAccessToken({ prompt: 'none' }); }
function handleAuthClick() { if (tokenClient) tokenClient.requestAccessToken({ prompt: 'consent' }); }
function handleSignoutClick() {
    const token = gapi.client.getToken();
    if (token !== null) {
        google.accounts.oauth2.revoke(token.access_token);
        gapi.client.setToken('');
        mainContainer.style.display = 'none';
        placeholder.style.display = 'block';
        authorizeButton.style.display = 'block';
        signoutButton.style.display = 'none';
    }
}


// --- Logic for "Review Queue" (formerly Inbox) ---
async function fetchEmailList() {
    showLoading('Loading review queue from Sheet...');
    try {
        if (QUEUE_SHEET_ID.includes('PASTE_YOUR')) { throw new Error('Queue Sheet ID is not configured'); }
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: QUEUE_SHEET_ID,
            range: 'Queue!A:E',
        });
        const rows = response.result.values;
        if (!rows || rows.length <= 1) {
            showError('No pending items found in the review queue.');
            return;
        }
        const headers = rows.shift().map(h => h.trim());
        const headerIndexes = {
            messageId: headers.indexOf('messageId'),
            subject: headers.indexOf('subject'),
            sender: headers.indexOf('sender'),
            receivedDate: headers.indexOf('receivedDate'),
            status: headers.indexOf('status')
        };
        if (Object.values(headerIndexes).includes(-1)) {
            throw new Error('Queue sheet is missing one or more required columns.');
        }
        const allEmails = rows.map((row, index) => ({
            rowIndex: index + 2,
            messageId: row[headerIndexes.messageId],
            subject: row[headerIndexes.subject],
            sender: row[headerIndexes.sender],
            receivedDate: row[headerIndexes.receivedDate],
            status: row[headerIndexes.status]
        }));
        
        // REVERTED: Directly filter and use the data without caching
        const pendingEmails = allEmails.filter(email => email.status === 'pending');
        hideLoading();
        
        if (pendingEmails.length > 0) {
            displayEmails(pendingEmails);
        } else {
            showError('No pending items found in the review queue.');
        }
    } catch (error) {
        console.error('Error fetching from Queue Sheet:', error);
        showError(`Failed to load review queue: ${error.result?.error?.message || error.message}`);
    }
}

function displayEmails(emails) {
    emailTableBody.innerHTML = '';
    emails.forEach(email => {
        const row = document.createElement('tr');
        row.setAttribute('data-message-id', email.messageId);
        row.innerHTML = `
            <td><span class="status-badge ${email.status.toLowerCase()}">${email.status}</span></td>
            <td>${email.sender || 'N/A'}</td>
            <td>${email.subject || '(No Subject)'}</td>
            <td>${email.receivedDate ? new Date(email.receivedDate).toLocaleString() : 'N/A'}</td>
            <td><button class="action-button">Process</button></td>
        `;
        const processButton = row.querySelector('.action-button');
        processButton.onclick = () => processSingleEmail(email, processButton);
        emailTableBody.appendChild(row);
    });
    emailTable.style.display = 'table';
}

async function processSingleEmail(email, button) {
    button.disabled = true;
    button.textContent = 'Processing...';
    button.classList.remove('error');
    try {
        if (PROCESSOR_WEBHOOK_URL.includes('PASTE_YOUR')) { throw new Error('Processor Webhook URL is not configured'); }
        const response = await fetch(PROCESSOR_WEBHOOK_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                messageId: email.messageId,
                queueRowIndex: email.rowIndex
            })
        });
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Processor workflow error: ${response.statusText} - ${errorText}`);
        }
        button.textContent = 'Success!';
        button.classList.add('success');
        showToast('Item processed successfully!', 'success');
        setTimeout(() => { fetchEmailList(); loadAllData(); }, 1500);
    } catch (error) {
        console.error('Error processing item:', error);
        button.disabled = false;
        button.textContent = 'Retry';
        button.classList.add('error');
        showToast(`Error: ${error.message}`, 'error');
    }
}

// BATCH PROCESSING FUNCTION REMOVED


// --- Dashboard & "Processed Queue" Data Loading ---
async function loadAllData() {
    try {
        const rangesToFetch = [
            ...SALES_TAB_NAMES.map(name => ({ sheetId: SALES_SHEET_ID, range: name })),
            ...PURCHASE_TAB_NAMES.map(name => ({ sheetId: PURCHASE_SHEET_ID, range: name }))
        ];
        const requests = rangesToFetch.map(req =>
            gapi.client.sheets.spreadsheets.values.get({
                spreadsheetId: req.sheetId,
                range: req.range
            })
        );
        const results = await Promise.allSettled(requests);
        allItemsCache = [];
        results.forEach((result, index) => {
            if (result.status === 'fulfilled' && result.value.result.values) {
                const cycle = rangesToFetch[index].sheetId === SALES_SHEET_ID ? 'Sales' : 'Purchase';
                allItemsCache.push(...processSheetData(result.value.result.values, cycle));
            }
        });
        filterAndRenderData('All');
    } catch (err) {
        console.error("Error loading data from dashboard sheets: ", err);
        showToast(`Error loading dashboard data. Check console.`, 'error');
    }
}

function processSheetData(data, cycle) {
    if (data.length < 2) return [];
    const headers = data.shift().map(h => h.toLowerCase().trim().replace(/ /g, '_'));
    return data.map((row, rowIndex) => {
        let item = { cycle: cycle, row_index: rowIndex + 2 };
        headers.forEach((header, i) => { item[header] = row[i]; });
        return item;
    });
}

function filterAndRenderData(cycle) {
    const filteredItems = (cycle === 'All') 
        ? allItemsCache 
        : allItemsCache.filter(item => item.cycle === cycle);

    const reviewQueueItems = filteredItems.filter(item => item.status === 'needs_review');
    const completedQueueItems = filteredItems.filter(item => item.status === 'completed');

    document.getElementById('kpi-total-docs').textContent = filteredItems.length;
    document.getElementById('kpi-review').textContent = reviewQueueItems.length;
    document.getElementById('kpi-completed').textContent = completedQueueItems.length;

    renderProcessedQueueTable(document.getElementById('dashboard-review-queue'), '<h3>Processed Queue (First 5)</h3>', reviewQueueItems.slice(0, 5));
    renderProcessedQueueTable(document.getElementById('full-review-queue-container'), '', reviewQueueItems);
}


// --- UI Rendering ---
function renderProcessedQueueTable(container, title, data) {
    container.innerHTML = title || '';
    if (data.length === 0) {
        container.innerHTML += `<p>No items in this queue.</p>`;
        return;
    }
    const table = document.createElement('table');
    table.className = 'data-table';
    const thead = table.createTHead();
    const headerRow = thead.insertRow();
    
    const headers = ['Doc Type', 'Cycle', 'Sender', 'Receive Date & Time', 'Action'];
    
    headers.forEach(text => {
        const th = document.createElement('th');
        th.textContent = text;
        headerRow.appendChild(th);
    });
    const tbody = table.createTBody();
    data.forEach(item => {
        const row = tbody.insertRow();
        renderProcessedQueueRow(row, item);
    });
    container.appendChild(table);
}

function renderProcessedQueueRow(row, item) {
    row.insertCell().textContent = item.document_type || 'N/A';
    row.insertCell().textContent = item.cycle || 'N/A';
    row.insertCell().textContent = item.partner_name || 'N/A';
    const dateValue = item.document_date;
    row.insertCell().textContent = dateValue ? new Date(dateValue).toLocaleString() : 'N/A';

    const actionCell = row.insertCell();
    actionCell.className = 'action-cell';
    const viewButton = document.createElement('button');
    viewButton.className = 'view-button';
    viewButton.textContent = 'View';
    viewButton.onclick = () => showViewModal(item);
    actionCell.appendChild(viewButton);
}


function showViewModal(item) {
    viewModalTitle.textContent = `Details for: ${item.document_type || 'N/A'} - ${item.reference_id || 'N/A'}`;
    let detailsHtml = '<div class="detail-grid">';
    for (const key in item) {
        if (item.hasOwnProperty(key)) {
            const label = key.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
            let value = item[key] || 'N/A';
            if ((key === 'line_items_json' || key === 'line_item_json') && value !== 'N/A') {
                try {
                    const items = JSON.parse(value).items || JSON.parse(value);
                    if (Array.isArray(items) && items.length > 0) {
                        let itemTable = '<table class="data-table"><thead><tr>';
                        const itemHeaders = Object.keys(items[0] || {});
                        itemHeaders.forEach(h => itemTable += `<th>${h.replace(/_/g, ' ')}</th>`);
                        itemTable += '</tr></thead><tbody>';
                        items.forEach(it => {
                            itemTable += '<tr>';
                            itemHeaders.forEach(h => itemTable += `<td>${it[h] || ''}</td>`);
                            itemTable += '</tr>';
                        });
                        itemTable += '</tbody></table>';
                        value = itemTable;
                    }
                } catch (e) { /* Fallback */ }
            }
            const tempDiv = document.createElement('div');
            tempDiv.textContent = value;
            value = tempDiv.innerHTML.replace(/\n/g, '<br>');
            
            detailsHtml += `<div class="detail-label">${label}:</div>`;
            detailsHtml += `<div class="detail-value">${value}</div>`;
        }
    }
    detailsHtml += '</div>';
    viewModalBody.innerHTML = detailsHtml;
    viewModal.style.display = 'flex';
}


// --- Shared Helper Functions ---
function formatCSVCell(text) {
    const str = String(text || '').trim();
    if (str.includes(',') || str.includes('"') || str.includes('\n')) {
        return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
}

function exportTableToCSV(containerId) {
    const container = document.getElementById(containerId);
    if (!container) {
        console.error(`Export failed: Container with id "${containerId}" not found.`);
        return;
    }

    const table = container.querySelector('table');
    if (!table) {
        showToast('No data available to export.', 'error');
        return;
    }

    const csv = [];
    const rows = table.querySelectorAll('tr');

    rows.forEach(row => {
        const rowData = [];
        const cells = row.querySelectorAll('th, td');
        
        cells.forEach((cell, index) => {
            const headerText = table.querySelector('th:nth-child(' + (index + 1) + ')').textContent.trim();
            if (headerText.toLowerCase() === 'action') {
                return;
            }
            rowData.push(formatCSVCell(cell.textContent));
        });
        csv.push(rowData.join(','));
    });

    const csvContent = csv.join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');

    const url = URL.createObjectURL(blob);
    const filename = `${containerId.replace(/-/g, '_')}_export.csv`;
    
    link.setAttribute('href', url);
    link.setAttribute('download', filename);
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showToast('Export successful!', 'success');
}

function showToast(message, type = 'success') {
    const toastContainer = document.getElementById('toast-container');
    if(!toastContainer) return;
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    toastContainer.appendChild(toast);
    setTimeout(() => { toast.remove(); }, 4000);
}

function showLoading(message) {
    loadingDiv.innerHTML = `<i class="fas fa-spinner fa-spin"></i> ${message}`;
    loadingDiv.style.display = 'block';
    errorDiv.style.display = 'none';
    emailTable.style.display = 'none';
}

function hideLoading() {
    loadingDiv.style.display = 'none';
}

function showError(message) {
    errorDiv.textContent = message;
    errorDiv.style.display = 'block';
    hideLoading();
}