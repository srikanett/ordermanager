// ===== ระบบจัดการออร์เดอร์ - JavaScript (Complete & Tested) =====

// ⚠️ IMPORTANT: เปลี่ยนเป็น URL ของคุณ
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbw3hclau08WNu6gGJ2Zzze-oY2wLoZ6mwUezborW4VeRrGF9kzQYnXFXMNIQxxvfPJJ/exec';
const APP_URL = window.location.origin + window.location.pathname;

// Global State
let allOrders = [];
let allCustomers = [];
let allProducts = [];
let allSheetNames = ['order'];
let currentOrderData = [];
let currentEditOrderId = null;
let currentEditCustomerId = null;
let currentEditProductId = null;

// DOM Elements
const statusOverlay = document.getElementById('status-overlay');
const progressTitle = document.getElementById('progress-title');
const progressText = document.getElementById('progress-text');
const adminPanel = document.getElementById('admin-panel');
const customerFormContainer = document.getElementById('customer-form-container');

// ===== UTILITIES =====

function showProgress(title = 'กำลังประมวลผล', text = 'โปรดรอสักครู่...') {
    progressTitle.textContent = title;
    progressText.textContent = text;
    statusOverlay.classList.remove('hidden');
}

function hideStatus() {
    statusOverlay.classList.add('hidden');
}

function showToast(icon, title) {
    const Toast = Swal.mixin({
        toast: true,
        position: 'top-end',
        showConfirmButton: false,
        timer: 3000,
        timerProgressBar: true,
        didOpen: (toast) => {
            toast.onmouseenter = Swal.stopTimer;
            toast.onmouseleave = Swal.resumeTimer;
        }
    });
    Toast.fire({ icon, title });
}

function formatBEDateTime(date = new Date()) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear() + 543;
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${day}-${month}-${year} : ${hours}:${minutes}`;
}

function generateOrderId() {
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = (now.getFullYear() + 543).toString().slice(-2);
    
    const prefix = `${year}${month}${day}`;
    // ✅ แปลง OrderID เป็น string ก่อน
    const todayOrders = allOrders.filter(o => o.OrderID && String(o.OrderID).startsWith(prefix));
    let maxSuffix = 0;
    todayOrders.forEach(o => {
        const suffix = parseInt(String(o.OrderID).slice(-3), 10);
        if (suffix > maxSuffix) maxSuffix = suffix;
    });
    const newSuffix = String(maxSuffix + 1).padStart(3, '0');
    return `${prefix}${newSuffix}`;
}

function formatPhoneForSheet(phone) {
    if (!phone) return '';
    phone = String(phone).trim();
    if (phone.startsWith("'")) return phone;
    if (phone.startsWith('0')) return `'${phone}`;
    return phone;
}

function formatCustomerName(name) {
    if (!name) return '';
    name = name.trim();
    if (name.startsWith('คุณ')) return name;
    return `คุณ${name}`;
}

function copyToClipboard(text) {
    if (navigator.clipboard) {
        navigator.clipboard.writeText(text).then(() => {
            showToast('success', 'คัดลอกแล้ว!');
        }, (err) => {
            showToast('error', 'คัดลอกไม่สำเร็จ');
        });
    }
}

function toBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result);
        reader.onerror = error => reject(error);
    });
}

async function uploadFileToDrive(file, folderId, fileName) {
    if (!file) return null;
    
    const base64 = await toBase64(file);
    const payload = {
        base64Data: base64,
        mimeType: file.type,
        folderId: folderId,
        fileName: fileName
    };
    
    const response = await gasFetch('uploadFile', payload);
    if (response && response.fileUrl) {
        return response.fileUrl;
    } else {
        throw new Error(response.error || 'Upload failed');
    }
}

// ===== GAS FETCH (FIXED - CORS Preflight 405) =====

async function gasFetch(action, payload) {
    try {
        console.log('📤 Sending:', action, payload);
        
        // ✅ ลบ headers ทั้งหมด - ป้องกัน CORS preflight 405 error
        const response = await fetch(SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action, payload })
        });
        
        console.log('📥 Response status:', response.status);
        
        const text = await response.text();
        console.log('📊 Response text:', text);
        
        if (!text) {
            throw new Error('Empty response from GAS');
        }
        
        const result = JSON.parse(text);
        
        if (!result.success) {
            throw new Error(result.error || 'Unknown GAS error');
        }
        
        return result.data;
        
    } catch (error) {
        console.error('❌ gasFetch Error:', action, error);
        showProgress('เกิดข้อผิดพลาด', error.message);
        setTimeout(hideStatus, 3000);
        throw error;
    }
}

// ===== INITIALIZATION =====

document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 App initialized');
    console.log('🔗 GAS URL:', SCRIPT_URL);
    checkUrlParams();
    setupEventListeners();
});

function checkUrlParams() {
    const urlParams = new URLSearchParams(window.location.search);
    const page = urlParams.get('page');
    const orderId = urlParams.get('orderId');

    if (page === 'customer-order' && orderId) {
        showCustomerForm(orderId);
    } else {
        showAdminPanel();
    }
}

function showAdminPanel() {
    adminPanel.hidden = false;
    customerFormContainer.hidden = true;
    fetchAllData();
}

async function showCustomerForm(orderId) {
    adminPanel.hidden = true;
    customerFormContainer.hidden = false;
    
    showProgress('กำลังโหลดข้อมูล...');
    try {
        const data = await gasFetch('getAllData', {});
        allCustomers = data.customers || [];
        
        const orderData = await gasFetch('getOrderById', { orderId });
        
        if (!orderData) {
            customerFormContainer.innerHTML = `<h2 class="text-2xl text-center text-red-400">ไม่พบข้อมูล ${orderId}</h2>`;
            return;
        }
        
        if (orderData.Status !== 'สร้างออร์เดอร์') {
            customerFormContainer.innerHTML = `
                <h1 class="text-3xl font-bold text-white mb-2 text-center">ออร์เดอร์ ${orderId}</h1>
                <p class="text-lg mb-6 text-center" style="color: var(--text-secondary);">ออร์เดอร์นี้ได้รับการยืนยันแล้ว</p>
                <div class="p-6 rounded-2xl text-center" style="background: rgba(50, 50, 46, 0.35);">
                    <h2 class="text-xl font-semibold mb-4" style="color: var(--status-success);">สถานะ: ${orderData.Status}</h2>
                    <p>ติดต่อแอดมิน หากมีข้อสงสัย</p>
                </div>`;
            return;
        }
        
        document.getElementById('customer-order-id').textContent = orderData.OrderID;
        
        document.getElementById('customer-item-list').innerHTML = '';
        for (let i = 1; i <= 5; i++) {
            const name = orderData[`Item${i}Name`];
            const price = orderData[`Item${i}Price`];
            if (name && price) {
                const itemEl = document.createElement('div');
                itemEl.className = 'flex justify-between items-center text-sm';
                itemEl.innerHTML = `
                    <span style="color: var(--text-secondary);">${name}</span>
                    <span>${parseFloat(price).toLocaleString()} ฿</span>
                `;
                document.getElementById('customer-item-list').appendChild(itemEl);
            }
        }
        
        document.getElementById('customer-total-price').textContent = `${parseFloat(orderData.TotalPrice).toLocaleString()} ฿`;
        
        setupCustomerAutocomplete();
        
    } catch (error) {
        console.error('Error loading form:', error);
    } finally {
        hideStatus();
    }
}

async function fetchAllData() {
    showProgress('กำลังโหลดข้อมูล...');
    try {
        const data = await gasFetch('getAllData', {});
        allOrders = data.orders || [];
        allCustomers = data.customers || [];
        allProducts = data.products || [];
        allSheetNames = ['order', ...(data.sheetNames || [])];
        
        console.log('✅ Data loaded:', { orders: allOrders.length, customers: allCustomers.length, products: allProducts.length });
        
        renderOrderTable();
        renderCustomerTable();
        renderProductTable();
        
        setupOrderAutocompletes();
        updateSheetNameDropdowns();
        
    } catch (error) {
        console.error('Failed to fetch data:', error);
        showToast('error', 'ไม่สามารถโหลดข้อมูล');
    } finally {
        hideStatus();
    }
}

function updateSheetNameDropdowns() {
    const extraSheets = allSheetNames.filter(name => !['order', 'รายชื่อ', 'สินค้า'].includes(name));
    
    const select = document.getElementById('print-sheet-select');
    select.innerHTML = '<option value="order">ชีต order</option>';
    extraSheets.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = `ชีต ${name}`;
        select.appendChild(option);
    });
}

// ===== EVENT LISTENERS =====

function setupEventListeners() {
    // Main Tabs
    document.getElementById('tab-btn-orders').addEventListener('click', () => {
        switchTab('order-mgmt-content', 'tab-btn-orders');
    });
    document.getElementById('tab-btn-data').addEventListener('click', () => {
        switchTab('data-mgmt-content', 'tab-btn-data');
    });
    
    // Order Sub-Tabs
    document.getElementById('order-subtab-btn-manage').addEventListener('click', () => {
        switchSubTab('order-subtab-manage', 'order-subtab-btn-manage');
        document.getElementById('order-manage-buttons').hidden = false;
    });
    document.getElementById('order-subtab-btn-print').addEventListener('click', () => {
        switchSubTab('order-subtab-print', 'order-subtab-btn-print');
        document.getElementById('order-manage-buttons').hidden = true;
        loadPrintTableData();
    });
    
    // Data Sub-Tabs
    document.getElementById('data-subtab-btn-customer').addEventListener('click', () => {
        switchSubTab('data-subtab-customer', 'data-subtab-btn-customer');
        document.getElementById('customer-manage-buttons').hidden = false;
        document.getElementById('product-manage-buttons').hidden = true;
    });
    document.getElementById('data-subtab-btn-product').addEventListener('click', () => {
        switchSubTab('data-subtab-product', 'data-subtab-btn-product');
        document.getElementById('customer-manage-buttons').hidden = true;
        document.getElementById('product-manage-buttons').hidden = false;
    });
    
    // Order Manage
    document.getElementById('order-search').addEventListener('input', renderOrderTable);
    document.getElementById('order-status-filter').addEventListener('change', renderOrderTable);
    document.getElementById('btn-add-order').addEventListener('click', () => showOrderModal(null));
    document.getElementById('order-select-all').addEventListener('change', toggleOrderSelectAll);
    document.getElementById('order-table-body').addEventListener('change', updateOrderBulkButtons);
    document.getElementById('btn-bulk-delete-order').addEventListener('click', bulkDeleteOrders);
    document.getElementById('btn-bulk-close-order').addEventListener('click', bulkCloseOrders);
    
    // Order Print
    document.getElementById('print-sheet-select').addEventListener('change', loadPrintTableData);
    document.getElementById('print-data-type-filter').addEventListener('change', renderPrintTable);
    document.getElementById('print-status-filter').addEventListener('change', renderPrintTable);
    document.getElementById('print-select-all').addEventListener('change', togglePrintSelectAll);
    document.getElementById('print-table-body').addEventListener('change', updatePrintBulkButtons);
    document.getElementById('btn-bulk-print').addEventListener('click', () => {
        const ids = getSelectedIds('print-table-body', 'print-checkbox');
        promptPrintChoice(ids);
    });
    document.getElementById('btn-label-settings').addEventListener('click', showLabelSettingsModal);
    
    // Customer
    document.getElementById('customer-search').addEventListener('input', renderCustomerTable);
    document.getElementById('btn-add-customer').addEventListener('click', () => showCustomerModal(null));
    document.getElementById('customer-select-all').addEventListener('change', toggleCustomerSelectAll);
    document.getElementById('customer-table-body').addEventListener('change', updateCustomerBulkButtons);
    document.getElementById('btn-bulk-delete-customer').addEventListener('click', bulkDeleteCustomers);
    
    // Product
    document.getElementById('product-search').addEventListener('input', renderProductTable);
    document.getElementById('btn-add-product').addEventListener('click', () => showProductModal(null));
    document.getElementById('product-select-all').addEventListener('change', toggleProductSelectAll);
    document.getElementById('product-table-body').addEventListener('change', updateProductBulkButtons);
    document.getElementById('btn-bulk-delete-product').addEventListener('click', bulkDeleteProducts);
    
    // Order Modal
    document.getElementById('btn-save-order').addEventListener('click', saveOrder);
    document.getElementById('btn-copy-link').addEventListener('click', () => {
        copyToClipboard(document.getElementById('order-customer-link').value);
    });
    document.getElementById('order-item-list').addEventListener('input', (e) => {
        if (e.target.classList.contains('order-item-name')) {
            const index = parseInt(e.target.closest('[data-item-index]').dataset.itemIndex);
            if (index < 5 && e.target.value.trim() !== '') {
                const nextItem = document.getElementById('order-item-list').querySelector(`[data-item-index="${index + 1}"]`);
                if (nextItem) nextItem.hidden = false;
            }
        }
        if (e.target.classList.contains('order-item-price')) {
            calculateOrderTotal();
        }
    });
    
    // Customer Modal
    document.getElementById('btn-save-customer').addEventListener('click', saveCustomer);
    
    // Product Modal
    document.getElementById('btn-save-product').addEventListener('click', saveProduct);
    document.getElementById('product-image-modal').addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                document.getElementById('product-preview').src = e.target.result;
                document.getElementById('product-preview').hidden = false;
                document.getElementById('btn-remove-product-image').hidden = false;
                document.getElementById('product-image-url-hidden').value = '';
            }
            reader.readAsDataURL(file);
        }
    });
    document.getElementById('btn-remove-product-image').addEventListener('click', () => {
        document.getElementById('product-image-modal').value = '';
        document.getElementById('product-preview').src = '';
        document.getElementById('product-preview').hidden = true;
        document.getElementById('btn-remove-product-image').hidden = true;
        document.getElementById('product-image-url-hidden').value = 'DELETED';
    });
    
    // View Modal
    document.getElementById('btn-copy-view-data').addEventListener('click', () => {
        const content = document.getElementById('view-modal-body').innerText || document.getElementById('view-modal-body').textContent;
        copyToClipboard(content);
    });
    document.getElementById('view-modal-body').addEventListener('click', (e) => {
        if (e.target.tagName === 'IMG' && e.target.dataset.viewable) {
            document.getElementById('image-view-modal-src').src = e.target.src;
            document.getElementById('image-view-modal').classList.remove('hidden');
        }
    });
    
    // Customer Form
    document.getElementById('customer-form').addEventListener('submit', submitCustomerOrder);
    document.getElementById('customer-slip').addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                document.getElementById('slip-preview').src = e.target.result;
                document.getElementById('slip-preview').classList.remove('hidden');
            }
            reader.readAsDataURL(file);
        } else {
            document.getElementById('slip-preview').classList.add('hidden');
        }
    });
}

// ===== TAB SWITCHING =====

function switchTab(contentId, tabId) {
    document.getElementById('order-mgmt-content').hidden = true;
    document.getElementById('data-mgmt-content').hidden = true;
    
    document.getElementById('tab-btn-orders').classList.remove('active');
    document.getElementById('tab-btn-data').classList.remove('active');
    
    document.getElementById(contentId).hidden = false;
    document.getElementById(tabId).classList.add('active');
}

function switchSubTab(contentId, tabId) {
    if (contentId.includes('order-subtab')) {
        document.getElementById('order-subtab-manage').hidden = true;
        document.getElementById('order-subtab-print').hidden = true;
        document.getElementById('order-subtab-btn-manage').classList.remove('active');
        document.getElementById('order-subtab-btn-print').classList.remove('active');
    } else if (contentId.includes('data-subtab')) {
        document.getElementById('data-subtab-customer').hidden = true;
        document.getElementById('data-subtab-product').hidden = true;
        document.getElementById('data-subtab-btn-customer').classList.remove('active');
        document.getElementById('data-subtab-btn-product').classList.remove('active');
    }
    
    document.getElementById(contentId).hidden = false;
    document.getElementById(tabId).classList.add('active');
}

// ===== RENDER ORDER TABLE =====

function renderOrderTable() {
    const searchTerm = document.getElementById('order-search').value.toLowerCase();
    const status = document.getElementById('order-status-filter').value;
    
    const filtered = allOrders.filter(order => {
        // ✅ แปลง OrderID เป็น string
        const matchesSearch = !searchTerm ||
            (order.OrderID && String(order.OrderID).includes(searchTerm)) ||
            (order.CustomerName && order.CustomerName.toLowerCase().includes(searchTerm)) ||
            (order.CustomerPhone && String(order.CustomerPhone).includes(searchTerm));
        
        const matchesStatus = (status === 'all') || (order.Status === status);
        return matchesSearch && matchesStatus;
    });
    
    const tbody = document.getElementById('order-table-body');
    const empty = document.getElementById('order-table-empty');
    
    if (filtered.length === 0) {
        tbody.innerHTML = '';
        empty.hidden = false;
        return;
    }
    
    empty.hidden = true;
    tbody.innerHTML = filtered.map((order, index) => {
        const rowClass = index % 2 === 0 ? 'table-row-light' : 'table-row-dark';
        const statusColors = {
            'สร้างออร์เดอร์': '#888',
            'รอตรวจสอบ': 'var(--status-error)',
            'รอจัดส่ง': 'var(--status-success)',
            'สำเร็จ': '#4CAF50'
        };
        const statusTextColor = (order.Status === 'รอจัดส่ง' || order.Status === 'สำเร็จ') ? '#333' : 'white';
        // ✅ แปลง OrderID เป็น string
        const orderId = String(order.OrderID);
        
        return `
            <tr class="${rowClass}">
                <td><input type="checkbox" class="styled-checkbox order-checkbox" data-id="${orderId}"></td>
                <td>${orderId}</td>
                <td>${order.CustomerName || ''}</td>
                <td>${order.TotalPrice ? parseFloat(order.TotalPrice).toLocaleString() : '0'}</td>
                <td><span class="px-2 py-1 rounded text-xs font-semibold" style="background: ${statusColors[order.Status] || '#888'}; color: ${statusTextColor};">${order.Status}</span></td>
                <td class="flex flex-wrap gap-2">
                    <button class="btn btn-secondary btn-icon btn-sm" onclick="showViewModal('order', '${orderId}')">👁️</button>
                    <button class="btn btn-secondary btn-icon btn-sm" onclick="showOrderModal('${orderId}')">✏️</button>
                    <button class="btn btn-error btn-icon btn-sm" onclick="deleteSingleItem('order', '${orderId}')">🗑️</button>
                </td>
            </tr>
        `;
    }).join('');
    
    updateOrderBulkButtons();
}

// ===== RENDER CUSTOMER TABLE (FIXED) =====

function renderCustomerTable() {
    const searchTerm = document.getElementById('customer-search').value.toLowerCase();
    
    const filtered = allCustomers.filter(customer => {
        return !searchTerm ||
            (customer.CustomerName && customer.CustomerName.toLowerCase().includes(searchTerm)) ||
            (customer.CustomerPhone && customer.CustomerPhone.includes(searchTerm)) ||
            (customer.CustomerID && customer.CustomerID.toString().includes(searchTerm));
    });
    
    const tbody = document.getElementById('customer-table-body');
    const empty = document.getElementById('customer-table-empty');
    
    if (filtered.length === 0) {
        tbody.innerHTML = '';
        empty.hidden = false;
        return;
    }
    
    empty.hidden = true;
    tbody.innerHTML = filtered.map((customer, index) => {
        const rowClass = index % 2 === 0 ? 'table-row-light' : 'table-row-dark';
        // ✅ แปลง CustomerID เป็น string
        const customerId = String(customer.CustomerID);
        
        return `
            <tr class="${rowClass}">
                <td><input type="checkbox" class="styled-checkbox customer-checkbox" data-id="${customerId}"></td>
                <td>${customerId}</td>
                <td>${customer.CustomerName || ''}</td>
                <td>${customer.CustomerPhone || ''}</td>
                <td>${customer.CustomerBirthday || ''}</td>
                <td class="flex flex-wrap gap-2">
                    <button class="btn btn-secondary btn-icon btn-sm" onclick="showViewModal('customer', '${customerId}')">👁️</button>
                    <button class="btn btn-secondary btn-icon btn-sm" onclick="showCustomerModal('${customerId}')">✏️</button>
                    <button class="btn btn-error btn-icon btn-sm" onclick="deleteSingleItem('รายชื่อ', '${customerId}')">🗑️</button>
                </td>
            </tr>
        `;
    }).join('');
    
    updateCustomerBulkButtons();
}

// ===== RENDER PRODUCT TABLE (FIXED) =====

function renderProductTable() {
    const searchTerm = document.getElementById('product-search').value.toLowerCase();
    
    const filtered = allProducts.filter(product => {
        return !searchTerm ||
            (product.ProductName && product.ProductName.toLowerCase().includes(searchTerm)) ||
            (product.ProductPrice && product.ProductPrice.toString().includes(searchTerm)) ||
            (product.ProductID && product.ProductID.toString().includes(searchTerm));
    });
    
    const tbody = document.getElementById('product-table-body');
    const empty = document.getElementById('product-table-empty');
    
    if (filtered.length === 0) {
        tbody.innerHTML = '';
        empty.hidden = false;
        return;
    }
    
    empty.hidden = true;
    tbody.innerHTML = filtered.map((product, index) => {
        const rowClass = index % 2 === 0 ? 'table-row-light' : 'table-row-dark';
        // ✅ แปลง ProductID เป็น string
        const productId = String(product.ProductID);
        
        return `
            <tr class="${rowClass}">
                <td><input type="checkbox" class="styled-checkbox product-checkbox" data-id="${productId}"></td>
                <td>${productId}</td>
                <td>${product.ProductName || ''}</td>
                <td>${product.ProductPrice ? parseFloat(product.ProductPrice).toLocaleString() : '0'} ฿</td>
                <td>${product.ImageUrl ? '<img src="' + product.ImageUrl + '" style="max-width:50px;max-height:50px;border-radius:4px;">' : '-'}</td>
                <td class="flex flex-wrap gap-2">
                    <button class="btn btn-secondary btn-icon btn-sm" onclick="showViewModal('product', '${productId}')">👁️</button>
                    <button class="btn btn-secondary btn-icon btn-sm" onclick="showProductModal('${productId}')">✏️</button>
                    <button class="btn btn-error btn-icon btn-sm" onclick="deleteSingleItem('สินค้า', '${productId}')">🗑️</button>
                </td>
            </tr>
        `;
    }).join('');
    
    updateProductBulkButtons();
}

// ===== LOAD PRINT TABLE =====

async function loadPrintTableData() {
    const sheetName = document.getElementById('print-sheet-select').value;
    showProgress('กำลังโหลด...');
    try {
        if (sheetName === 'order') {
            currentOrderData = [...allOrders];
        } else {
            const data = await gasFetch('getSheetData', { sheetName });
            currentOrderData = data.orders || [];
        }
        renderPrintTable();
    } catch (error) {
        console.error('Error:', error);
    } finally {
        hideStatus();
    }
}

function renderPrintTable() {
    const dataType = document.getElementById('print-data-type-filter').value;
    const status = document.getElementById('print-status-filter').value;
    
    const filtered = currentOrderData.filter(item => {
        // ✅ แปลง OrderID เป็น string
        const matchesType = (dataType === 'all') || (dataType === 'orderOnly' && item.OrderID);
        const matchesStatus = (status === 'all') || (item.Status === status);
        return matchesType && matchesStatus;
    });
    
    const tbody = document.getElementById('print-table-body');
    const empty = document.getElementById('print-table-empty');
    
    if (filtered.length === 0) {
        tbody.innerHTML = '';
        empty.hidden = false;
        return;
    }
    
    empty.hidden = true;
    tbody.innerHTML = filtered.map((item, i) => {
        const id = item.OrderID || item.No;
        const type = item.OrderID ? 'ออร์เดอร์' : 'ทั่วไป';
        
        return `
            <tr class="${i % 2 === 0 ? 'table-row-light' : 'table-row-dark'}">
                <td><input type="checkbox" class="styled-checkbox print-checkbox" data-id="${String(id)}"></td>
                <td>${id}</td>
                <td>${item.CustomerName}</td>
                <td>${parseFloat(item.TotalPrice).toLocaleString()}</td>
                <td><span class="px-2 py-1 rounded text-xs">${type}</span></td>
                <td class="flex gap-2">
                    <button class="btn btn-secondary btn-icon btn-sm" onclick="showViewModal('printItem', '${String(id)}')">👁️</button>
                    <button class="btn btn-secondary btn-icon btn-sm" onclick="promptPrintChoice(['${String(id)}'])">🖨️</button>
                </td>
            </tr>
        `;
    }).join('');
    updatePrintBulkButtons();
}

// ===== AUTOCOMPLETE =====

function setupOrderAutocompletes() {
    setupAutocomplete(
        document.getElementById('order-customer-name'),
        () => allCustomers,
        (item) => `<div><strong>${item.CustomerName}</strong><div style="font-size:0.8em">${item.CustomerPhone}</div></div>`,
        (item) => {
            document.getElementById('order-customer-name').value = item.CustomerName;
            document.getElementById('order-customer-address').value = item.CustomerAddress;
            document.getElementById('order-customer-phone').value = String(item.CustomerPhone).replace("'", "");
        }
    );
    
    document.querySelectorAll('.order-item-name').forEach(input => {
        setupAutocomplete(
            input,
            () => allProducts,
            (item) => `<div><strong>${item.ProductName}</strong><div style="font-size:0.8em">${parseFloat(item.ProductPrice).toLocaleString()} ฿</div></div>`,
            (item) => {
                input.value = item.ProductName;
                const priceInput = input.closest('.grid').querySelector('.order-item-price');
                priceInput.value = item.ProductPrice;
                calculateOrderTotal();
            }
        );
    });
}

function setupCustomerAutocomplete() {
    setupAutocomplete(
        document.getElementById('customer-name'),
        () => allCustomers,
        (item) => `<div><strong>${item.CustomerName}</strong><div style="font-size:0.8em">${item.CustomerPhone}</div></div>`,
        (item) => {
            document.getElementById('customer-name').value = item.CustomerName;
            document.getElementById('customer-address').value = item.CustomerAddress;
            document.getElementById('customer-phone').value = String(item.CustomerPhone).replace("'", "");
        }
    );
}

function setupAutocomplete(inp, dataCallback, renderCallback, selectCallback) {
    let currentFocus;
    
    inp.addEventListener("input", function() {
        let val = this.value;
        closeAllLists();
        if (!val) return;
        
        currentFocus = -1;
        const data = dataCallback();
        if (!data) return;
        
        let a = document.createElement("DIV");
        a.setAttribute("class", "autocomplete-items");
        this.parentNode.appendChild(a);
        
        let count = 0;
        data.forEach((item, i) => {
            let match = Object.values(item).some(v => String(v).toLowerCase().includes(val.toLowerCase()));
            
            if (match && count < 10) {
                count++;
                let b = document.createElement("DIV");
                b.innerHTML = renderCallback(item);
                b.addEventListener("click", () => {
                    selectCallback(item);
                    closeAllLists();
                });
                a.appendChild(b);
            }
        });
    });
    
    inp.addEventListener("keydown", function(e) {
        let x = this.parentNode.querySelector(".autocomplete-items");
        if (x) x = x.getElementsByTagName("div");
        if (e.keyCode == 40) {
            currentFocus++;
            addActive(x);
        } else if (e.keyCode == 38) {
            currentFocus--;
            addActive(x);
        } else if (e.keyCode == 13) {
            e.preventDefault();
            if (currentFocus > -1 && x) x[currentFocus].click();
        }
    });
    
    function addActive(x) {
        if (!x) return;
        x.forEach(item => item.classList.remove("autocomplete-active"));
        if (currentFocus >= x.length) currentFocus = 0;
        if (currentFocus < 0) currentFocus = x.length - 1;
        if (x[currentFocus]) x[currentFocus].classList.add("autocomplete-active");
    }
    
    function closeAllLists(elmnt) {
        let x = document.getElementsByClassName("autocomplete-items");
        for (let i = x.length - 1; i >= 0; i--) {
            if (elmnt != x[i] && elmnt != inp) {
                x[i].parentNode.removeChild(x[i]);
            }
        }
    }
    
    document.addEventListener("click", (e) => closeAllLists(e.target));
}

function calculateOrderTotal() {
    let total = 0;
    document.querySelectorAll('.order-item-price').forEach(input => {
        const price = parseFloat(input.value);
        if (!isNaN(price)) total += price;
    });
    document.getElementById('order-total-price').value = total > 0 ? total : '';
}

// ===== MODAL FUNCTIONS =====

async function showOrderModal(orderId) {
    currentEditOrderId = orderId;
    document.getElementById('order-form').reset();
    document.querySelectorAll('[data-item-index]').forEach((el, i) => {
        el.hidden = (i > 0);
    });
    
    if (orderId) {
        document.getElementById('order-modal-title').textContent = `แก้ไข: ${orderId}`;
        // ✅ แปลง orderId เป็น string เพื่อเทียบ
        const order = allOrders.find(o => String(o.OrderID) === String(orderId));
        if (!order) {
            showToast('error', 'ไม่พบออร์เดอร์');
            return;
        }
        
        document.getElementById('order-id').value = order.OrderID;
        document.getElementById('order-id-display').value = order.OrderID;
        document.getElementById('order-date').value = order.Date;
        document.getElementById('order-customer-name').value = order.CustomerName;
        document.getElementById('order-customer-address').value = order.CustomerAddress;
        document.getElementById('order-customer-phone').value = order.CustomerPhone ? String(order.CustomerPhone).replace("'", "") : "";
        
        for (let i = 1; i <= 5; i++) {
            const nameEl = document.querySelector(`[data-item-index="${i}"] .order-item-name`);
            const priceEl = document.querySelector(`[data-item-index="${i}"] .order-item-price`);
            
            nameEl.value = order[`Item${i}Name`] || '';
            priceEl.value = order[`Item${i}Price`] || '';
            
            if (order[`Item${i}Name`]) {
                document.querySelector(`[data-item-index="${i}"]`).hidden = false;
            }
        }
        
        document.getElementById('order-total-price').value = order.TotalPrice;
        document.getElementById('order-status').value = order.Status;
        
    } else {
        document.getElementById('order-modal-title').textContent = 'เพิ่มออร์เดอร์';
        const newId = generateOrderId();
        document.getElementById('order-id-display').value = newId;
        document.getElementById('order-date').value = formatBEDateTime();
        document.getElementById('order-status').value = 'สร้างออร์เดอร์';
    }
    
    const idForLink = orderId || document.getElementById('order-id-display').value;
    const link = `${APP_URL}?page=customer-order&orderId=${idForLink}`;
    document.getElementById('order-customer-link').value = link;
    
    setupOrderAutocompletes();
    document.getElementById('order-modal').classList.remove('hidden');
}

function showCustomerModal(customerId) {
    currentEditCustomerId = customerId;
    document.getElementById('customer-form-modal').reset();
    
    if (customerId) {
        document.getElementById('customer-modal-title').textContent = `แก้ไข: ${customerId}`;
        // ✅ แปลง customerId เป็น string เพื่อเทียบ
        const customer = allCustomers.find(c => String(c.CustomerID) === String(customerId));
        if (!customer) {
            showToast('error', 'ไม่พบลูกค้า');
            return;
        }
        document.getElementById('customer-id').value = customer.CustomerID;
        document.getElementById('customer-name-modal').value = customer.CustomerName;
        document.getElementById('customer-address-modal').value = customer.CustomerAddress;
        document.getElementById('customer-phone-modal').value = customer.CustomerPhone ? String(customer.CustomerPhone).replace("'", "") : "";
        document.getElementById('customer-birthday-modal').value = customer.CustomerBirthday;
    } else {
        document.getElementById('customer-modal-title').textContent = 'เพิ่มลูกค้า';
        document.getElementById('customer-id').value = '';
    }
    
    document.getElementById('customer-modal').classList.remove('hidden');
}

function showProductModal(productId) {
    currentEditProductId = productId;
    document.getElementById('product-form-modal').reset();
    document.getElementById('product-preview').src = '';
    document.getElementById('product-preview').hidden = true;
    document.getElementById('btn-remove-product-image').hidden = true;
    document.getElementById('product-image-url-hidden').value = '';
    
    if (productId) {
        document.getElementById('product-modal-title').textContent = `แก้ไข: ${productId}`;
        // ✅ แปลง productId เป็น string เพื่อเทียบ
        const product = allProducts.find(p => String(p.ProductID) === String(productId));
        if (!product) {
            showToast('error', 'ไม่พบสินค้า');
            return;
        }
        document.getElementById('product-id').value = product.ProductID;
        document.getElementById('product-name-modal').value = product.ProductName;
        document.getElementById('product-price-modal').value = product.ProductPrice;
        if (product.ImageUrl) {
            document.getElementById('product-preview').src = product.ImageUrl;
            document.getElementById('product-preview').hidden = false;
            document.getElementById('btn-remove-product-image').hidden = false;
            document.getElementById('product-image-url-hidden').value = product.ImageUrl;
        }
    } else {
        document.getElementById('product-modal-title').textContent = 'เพิ่มสินค้า';
        document.getElementById('product-id').value = '';
    }
    
    document.getElementById('product-modal').classList.remove('hidden');
}

function showViewModal(type, id) {
    let data, title, content = '';
    
    if (type === 'order') {
        // ✅ แปลง id เป็น string เพื่อเทียบ
        data = allOrders.find(o => String(o.OrderID) === String(id));
        title = `ออร์เดอร์: ${id}`;
        if (data) {
            content = `
                <p><strong>OrderID:</strong> ${data.OrderID}</p>
                <p><strong>วันที่:</strong> ${data.Date}</p>
                <p><strong>ชื่อ:</strong> ${data.CustomerName}</p>
                <p><strong>ที่อยู่:</strong> ${data.CustomerAddress}</p>
                <p><strong>เบอร์:</strong> ${data.CustomerPhone}</p>
                <hr style="border-color: rgba(255,255,255,0.2)">
                <h4><strong>รายการ:</strong></h4>
                ${[1,2,3,4,5].map(i => {
                    const name = data[`Item${i}Name`];
                    const price = data[`Item${i}Price`];
                    return (name && price) ? `<p>${i}. ${name} - ${parseFloat(price).toLocaleString()} ฿</p>` : '';
                }).join('')}
                <hr style="border-color: rgba(255,255,255,0.2)">
                <p><strong>รวม:</strong> ${parseFloat(data.TotalPrice).toLocaleString()} ฿</p>
                <p><strong>สถานะ:</strong> ${data.Status}</p>
                ${data.SlipURL ? `<p><strong>สลิป:</strong> <img src="${data.SlipURL}" data-viewable="true" class="rounded max-w-sm mt-2 cursor-pointer"></p>` : ''}
            `;
        }
    } else if (type === 'customer') {
        data = allCustomers.find(c => String(c.CustomerID) === String(id));
        title = `ลูกค้า: ${id}`;
        if (data) {
            content = `
                <p><strong>ID:</strong> ${data.CustomerID}</p>
                <p><strong>ชื่อ:</strong> ${data.CustomerName}</p>
                <p><strong>ที่อยู่:</strong> ${data.CustomerAddress}</p>
                <p><strong>เบอร์:</strong> ${data.CustomerPhone}</p>
                <p><strong>วันเกิด:</strong> ${data.CustomerBirthday}</p>
            `;
        }
    } else if (type === 'product') {
        data = allProducts.find(p => String(p.ProductID) === String(id));
        title = `สินค้า: ${id}`;
        if (data) {
            content = `
                <p><strong>ID:</strong> ${data.ProductID}</p>
                <p><strong>ชื่อ:</strong> ${data.ProductName}</p>
                <p><strong>ราคา:</strong> ${parseFloat(data.ProductPrice).toLocaleString()} ฿</p>
                ${data.ImageUrl ? `<p><strong>รูป:</strong> <img src="${data.ImageUrl}" data-viewable="true" class="rounded max-w-sm mt-2 cursor-pointer"></p>` : ''}
            `;
        }
    }
    
    document.getElementById('view-modal-title').textContent = title;
    document.getElementById('view-modal-body').innerHTML = content;
    document.getElementById('view-modal').classList.remove('hidden');
}

// ===== SAVE FUNCTIONS =====

async function saveOrder() {
    showProgress('กำลังบันทึก...');
    
    const itemData = {};
    for (let i = 1; i <= 5; i++) {
        const nameInput = document.querySelector(`[data-item-index="${i}"] .order-item-name`);
        const priceInput = document.querySelector(`[data-item-index="${i}"] .order-item-price`);
        itemData[`Item${i}Name`] = nameInput.value;
        itemData[`Item${i}Price`] = priceInput.value;
    }
    
    const payload = {
        OrderID: document.getElementById('order-id').value || document.getElementById('order-id-display').value,
        Date: document.getElementById('order-date').value,
        CustomerName: formatCustomerName(document.getElementById('order-customer-name').value),
        CustomerAddress: document.getElementById('order-customer-address').value,
        CustomerPhone: formatPhoneForSheet(document.getElementById('order-customer-phone').value),
        ...itemData,
        TotalPrice: document.getElementById('order-total-price').value,
        Status: document.getElementById('order-status').value,
        SlipURL: ''
    };
    
    const action = currentEditOrderId ? 'updateOrder' : 'createOrder';
    
    try {
        await gasFetch(action, payload);
        await fetchAllData();
        document.getElementById('order-modal').classList.add('hidden');
        showToast('success', 'บันทึกสำเร็จ!');
    } catch (error) {
        console.error('Save failed:', error);
    } finally {
        hideStatus();
    }
}

async function saveCustomer() {
    showProgress('กำลังบันทึก...');
    
    const payload = {
        CustomerID: document.getElementById('customer-id').value,
        CustomerName: formatCustomerName(document.getElementById('customer-name-modal').value),
        CustomerAddress: document.getElementById('customer-address-modal').value,
        CustomerPhone: formatPhoneForSheet(document.getElementById('customer-phone-modal').value),
        CustomerBirthday: document.getElementById('customer-birthday-modal').value,
    };
    
    const action = currentEditCustomerId ? 'updateCustomer' : 'createCustomer';
    
    try {
        await gasFetch(action, payload);
        await fetchAllData();
        document.getElementById('customer-modal').classList.add('hidden');
        showToast('success', 'บันทึกสำเร็จ!');
    } catch (error) {
        console.error('Save failed:', error);
    } finally {
        hideStatus();
    }
}

async function saveProduct() {
    showProgress('กำลังบันทึก...');
    
    const file = document.getElementById('product-image-modal').files[0];
    let imageUrl = document.getElementById('product-image-url-hidden').value;
    
    try {
        if (file) {
            const fileName = `${document.getElementById('product-name-modal').value}`;
            imageUrl = await uploadFileToDrive(file, '1bmmwfhO6718C3LcAnYvhoA8vB9SbKNcq', fileName);
        } else if (imageUrl === 'DELETED') {
            imageUrl = '';
        }
        
        const payload = {
            ProductID: document.getElementById('product-id').value,
            ProductName: document.getElementById('product-name-modal').value,
            ProductPrice: document.getElementById('product-price-modal').value,
            ImageUrl: imageUrl
        };
        
        const action = currentEditProductId ? 'updateProduct' : 'createProduct';
        await gasFetch(action, payload);
        
        await fetchAllData();
        document.getElementById('product-modal').classList.add('hidden');
        showToast('success', 'บันทึกสำเร็จ!');
        
    } catch (error) {
        console.error('Save failed:', error);
    } finally {
        hideStatus();
    }
}

async function submitCustomerOrder(e) {
    e.preventDefault();
    const orderId = document.getElementById('customer-order-id').textContent;
    const file = document.getElementById('customer-slip').files[0];
    
    if (!file) {
        showToast('error', 'กรุณาอัปโหลดสลิป');
        return;
    }
    
    showProgress('กำลังยืนยัน...');
    
    try {
        const customerName = formatCustomerName(document.getElementById('customer-name').value);
        const total = document.getElementById('customer-total-price').textContent.replace(' ฿', '');
        const fileName = `${customerName} ${total}`;
        
        const slipUrl = await uploadFileToDrive(file, '1xXUofsUJqYp2NfSlDTFcSugvAnY0rAi0', fileName);
        
        if (!slipUrl) throw new Error('ไม่สามารถอัปโหลดได้');
        
        const payload = {
            OrderID: orderId,
            CustomerName: customerName,
            CustomerAddress: document.getElementById('customer-address').value,
            CustomerPhone: formatPhoneForSheet(document.getElementById('customer-phone').value),
            SlipURL: slipUrl,
            Status: 'รอตรวจสอบ'
        };
        
        await gasFetch('updateCustomerOrder', payload);
        
        hideStatus();
        customerFormContainer.innerHTML = `
            <div class="p-6 text-center">
                <h1 class="text-3xl font-bold text-white mb-2">ยืนยันสำเร็จ!</h1>
                <p class="text-lg mb-6" style="color: var(--text-secondary);">ออร์เดอร์ ${orderId} ส่งให้แอดมิน</p>
                <div style="color: var(--status-success); font-size: 5rem;">✔</div>
            </div>
        `;
        
    } catch (error) {
        console.error('Submit failed:', error);
        showProgress('เกิดข้อผิดพลาด', error.message);
        setTimeout(hideStatus, 3000);
    }
}

// ===== DELETE FUNCTIONS =====

function getSelectedIds(tbodyId, checkboxClass) {
    const checkboxes = document.getElementById(tbodyId).querySelectorAll(`.${checkboxClass}:checked`);
    // ✅ แปลง id เป็น string
    return Array.from(checkboxes).map(cb => String(cb.dataset.id));
}

async function deleteSingleItem(sheetName, id) {
    // ✅ แปลง id เป็น string
    const strId = String(id);
    const result = await Swal.fire({
        title: 'แน่ใจ?',
        text: `ลบ ${strId}`,
        icon: 'warning',
        showCancelButton: true,
        confirmButtonText: 'ลบ',
        customClass: { popup: 'swal2-popup', confirmButton: 'swal2-confirm', cancelButton: 'swal2-cancel' }
    });
    
    if (result.isConfirmed) {
        performDelete(sheetName, [strId]);
    }
}

async function bulkDeleteOrders() {
    const ids = getSelectedIds('order-table-body', 'order-checkbox');
    const result = await Swal.fire({
        title: 'ลบ ' + ids.length + ' ออร์เดอร์?',
        icon: 'warning',
        showCancelButton: true,
        confirmButtonText: 'ลบ',
        customClass: { popup: 'swal2-popup', confirmButton: 'swal2-confirm', cancelButton: 'swal2-cancel' }
    });
    if (result.isConfirmed) performDelete('order', ids);
}

async function bulkDeleteCustomers() {
    const ids = getSelectedIds('customer-table-body', 'customer-checkbox');
    const result = await Swal.fire({
        title: 'ลบ ' + ids.length + ' ลูกค้า?',
        icon: 'warning',
        showCancelButton: true,
        confirmButtonText: 'ลบ',
        customClass: { popup: 'swal2-popup', confirmButton: 'swal2-confirm', cancelButton: 'swal2-cancel' }
    });
    if (result.isConfirmed) performDelete('รายชื่อ', ids);
}

async function bulkDeleteProducts() {
    const ids = getSelectedIds('product-table-body', 'product-checkbox');
    const result = await Swal.fire({
        title: 'ลบ ' + ids.length + ' สินค้า?',
        icon: 'warning',
        showCancelButton: true,
        confirmButtonText: 'ลบ',
        customClass: { popup: 'swal2-popup', confirmButton: 'swal2-confirm', cancelButton: 'swal2-cancel' }
    });
    if (result.isConfirmed) performDelete('สินค้า', ids);
}

async function performDelete(sheetName, ids) {
    showProgress('กำลังลบ...');
    try {
        await gasFetch('deleteItems', { sheetName, ids });
        await fetchAllData();
        showToast('success', 'ลบสำเร็จ');
    } catch (error) {
        console.error('Delete failed:', error);
    } finally {
        hideStatus();
    }
}

async function bulkCloseOrders() {
    const ids = getSelectedIds('order-table-body', 'order-checkbox');
    const { value: targetSheet } = await Swal.fire({
        title: 'ปิดงาน ' + ids.length + ' ออร์เดอร์',
        input: 'select',
        inputOptions: allSheetNames
            .filter(name => !['order', 'รายชื่อ', 'สินค้า'].includes(name))
            .reduce((acc, name) => ({ ...acc, [name]: name }), {}),
        showCancelButton: true,
        confirmButtonText: 'ย้าย',
        customClass: { popup: 'swal2-popup', confirmButton: 'swal2-confirm', cancelButton: 'swal2-cancel' }
    });

    if (targetSheet) {
        showProgress('กำลังปิดงาน...');
        try {
            await gasFetch('moveOrdersToSheet', { ids, targetSheetName: targetSheet });
            await fetchAllData();
            showToast('success', 'ปิดงานสำเร็จ');
        } catch (error) {
            console.error('Close failed:', error);
        } finally {
            hideStatus();
        }
    }
}

// ===== BULK SELECT =====

function toggleOrderSelectAll() {
    const selectAll = document.getElementById('order-select-all').checked;
    document.querySelectorAll('.order-checkbox').forEach(cb => cb.checked = selectAll);
    updateOrderBulkButtons();
}

function toggleCustomerSelectAll() {
    const selectAll = document.getElementById('customer-select-all').checked;
    document.querySelectorAll('.customer-checkbox').forEach(cb => cb.checked = selectAll);
    updateCustomerBulkButtons();
}

function toggleProductSelectAll() {
    const selectAll = document.getElementById('product-select-all').checked;
    document.querySelectorAll('.product-checkbox').forEach(cb => cb.checked = selectAll);
    updateProductBulkButtons();
}

function togglePrintSelectAll() {
    const selectAll = document.getElementById('print-select-all').checked;
    document.querySelectorAll('.print-checkbox').forEach(cb => cb.checked = selectAll);
    updatePrintBulkButtons();
}

function updateOrderBulkButtons() {
    const count = getSelectedIds('order-table-body', 'order-checkbox').length;
    document.getElementById('btn-bulk-delete-order').hidden = (count === 0);
    document.getElementById('btn-bulk-delete-order').textContent = `ลบ (${count})`;
}

function updateCustomerBulkButtons() {
    const count = getSelectedIds('customer-table-body', 'customer-checkbox').length;
    document.getElementById('btn-bulk-delete-customer').hidden = (count === 0);
    document.getElementById('btn-bulk-delete-customer').textContent = `ลบ (${count})`;
}

function updateProductBulkButtons() {
    const count = getSelectedIds('product-table-body', 'product-checkbox').length;
    document.getElementById('btn-bulk-delete-product').hidden = (count === 0);
    document.getElementById('btn-bulk-delete-product').textContent = `ลบ (${count})`;
}

function updatePrintBulkButtons() {
    const count = getSelectedIds('print-table-body', 'print-checkbox').length;
    document.getElementById('btn-bulk-print').disabled = (count === 0);
    document.getElementById('btn-bulk-print').textContent = `พิมพ์ (${count})`;
}

// ===== PRINT FUNCTIONS =====

function promptPrintChoice(ids) {
    Swal.fire({
        title: 'เลือกขนาด',
        icon: 'question',
        showCancelButton: true,
        confirmButtonText: 'Label 100x75',
        cancelButtonText: 'A6',
        customClass: { popup: 'swal2-popup', confirmButton: 'swal2-confirm', cancelButton: 'swal2-cancel' }
    }).then((result) => {
        if (result.isConfirmed) {
            printLabels(ids, '100x75');
        } else if (result.dismiss === Swal.DismissReason.cancel) {
            printLabels(ids, 'A6');
        }
    });
}

function printLabels(ids, format) {
    // ✅ แปลง ids เป็น string ก่อนค้นหา
    const ordersToPrint = ids.map(id => 
        currentOrderData.find(o => String(o.OrderID) === String(id) || String(o.No) === String(id))
    ).filter(Boolean);
    
    let htmlContent = format === 'A6' ? generateA6Html(ordersToPrint) : generate100x75Html(ordersToPrint);
    showPrintPreviewModal(htmlContent, format);
}

function showPrintPreviewModal(htmlContent, format) {
    Swal.fire({
        title: `ตัวอย่าง ${format}`,
        html: `<iframe srcdoc='${escape(htmlContent)}' style='width:100%;height:60vh;border:1px solid #ccc;border-radius:0.5rem;'></iframe>`,
        width: '90vw',
        showCancelButton: true,
        confirmButtonText: 'พิมพ์',
        customClass: { popup: 'swal2-popup', confirmButton: 'swal2-confirm', cancelButton: 'swal2-cancel' }
    }).then((result) => {
        if (result.isConfirmed) {
            const iframe = document.querySelector('iframe');
            if (iframe && iframe.contentWindow) {
                iframe.contentWindow.focus();
                iframe.contentWindow.print();
            }
        }
    });
}

function generateA6Html(orders) {
    return `
    <html>
    <head>
        <title>Print A6</title>
        <meta charset="utf-8">
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Thai&display=swap');
            body { font-family: 'Noto Sans Thai'; margin: 0; }
            @page { size: A6 landscape; margin: 1cm; }
            .page {
                width: 148mm; height: 105mm; padding: 1cm; box-sizing: border-box;
                border: 1px dashed #ccc; display: flex; flex-direction: column;
                justify-content: center; page-break-after: always;
            }
            h1 { margin: 0 0 5mm 0; font-size: 18px; }
            .name { font-size: 16px; font-weight: bold; margin-bottom: 3mm; }
            .address { font-size: 14px; margin-bottom: 3mm; }
            .phone { font-size: 16px; font-weight: bold; }
        </style>
    </head>
    <body>
        ${orders.map(o => `
        <div class="page">
            <h1>ผู้รับ (${o.OrderID})</h1>
            <div class="name">${o.CustomerName}</div>
            <div class="address">${o.CustomerAddress}</div>
            <div class="phone">โทร. ${o.CustomerPhone}</div>
        </div>
        `).join('')}
    </body>
    </html>
    `;
}

function generate100x75Html(orders) {
    return `
    <html>
    <head>
        <title>Print 100x75</title>
        <meta charset="utf-8">
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Thai&display=swap');
            body { font-family: 'Noto Sans Thai'; margin: 0; }
            @page { size: 100mm 75mm; margin: 5mm; }
            .page {
                width: 100mm; height: 75mm; padding: 5mm; box-sizing: border-box;
                border: 1px dashed #ccc; display: flex; flex-direction: column;
                justify-content: center; page-break-after: always;
            }
            h1 { margin: 0 0 3mm 0; font-size: 16px; }
            .name { font-size: 14px; font-weight: bold; margin-bottom: 2mm; }
            .address { font-size: 12px; margin-bottom: 2mm; }
            .phone { font-size: 14px; font-weight: bold; }
        </style>
    </head>
    <body>
        ${orders.map(o => `
        <div class="page">
            <h1>ผู้รับ (${o.OrderID})</h1>
            <div class="name">${o.CustomerName}</div>
            <div class="address">${o.CustomerAddress}</div>
            <div class="phone">โทร. ${o.CustomerPhone}</div>
        </div>
        `).join('')}
    </body>
    </html>
    `;
}

async function showLabelSettingsModal() {
    Swal.fire({
        title: 'ตั้งค่าฉลาก',
        html: `<p>ตั้งค่าจะเก็บใน localStorage ของเบราว์เซอร์</p>`,
        icon: 'info',
        confirmButtonText: 'ปิด',
        customClass: { popup: 'swal2-popup', confirmButton: 'swal2-confirm' }
    });
}

console.log('✅ Script loaded successfully');
