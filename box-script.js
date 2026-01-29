// Google Sheets Configuration
const SHEET_ID = '1jT55HXOqVATpFRmUmtm8GrraipfvNjNDKzlQXldER_M';
const SHEET_NAME = 'Box';
const API_KEY = 'AIzaSyBSsHfiNlbGZYLy_kfVJCNHJtQVVyxmxKU'; // You'll need to replace this with your API key

// Global variables
let allData = [];
let filteredData = [];
let sortColumn = null;
let sortDirection = 'asc';

// DOM Elements
const searchInput = document.getElementById('searchInput');
const shelfFilter = document.getElementById('shelfFilter');
const tableBody = document.getElementById('tableBody');
const headerCheckbox = document.getElementById('headerCheckbox');
const selectAllBtn = document.getElementById('selectAllBtn');
const deselectAllBtn = document.getElementById('deselectAllBtn');
const printSelectedBtn = document.getElementById('printSelectedBtn');
const totalItemsEl = document.getElementById('totalItems');
const selectedItemsEl = document.getElementById('selectedItems');
const filteredItemsEl = document.getElementById('filteredItems');

// Initialize the application
async function init() {
    try {
        await loadData();
        setupEventListeners();
        populateShelfFilter();
        renderTable();
        updateStats();
    } catch (error) {
        console.error('Error initializing app:', error);
        showError('ไม่สามารถโหลดข้อมูลได้ กรุณาตรวจสอบการเชื่อมต่อ');
    }
}

// Load data from Google Sheets
async function loadData() {
    const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${SHEET_NAME}`;

    try {
        const response = await fetch(url);
        const text = await response.text();
        const json = JSON.parse(text.substring(47).slice(0, -2));

        const rows = json.table.rows;
        allData = rows.map((row, index) => ({
            id: index,
            box: row.c[0]?.v || '',
            shelf: row.c[1]?.v || '',
            selected: false
        }));

        filteredData = [...allData];
    } catch (error) {
        console.error('Error loading data:', error);
        throw error;
    }
}

// Populate shelf filter dropdown
function populateShelfFilter() {
    const shelves = [...new Set(allData.map(item => item.shelf))].filter(Boolean).sort();

    shelfFilter.innerHTML = '<option value="">ทั้งหมด</option>';
    shelves.forEach(shelf => {
        const option = document.createElement('option');
        option.value = shelf;
        option.textContent = shelf;
        shelfFilter.appendChild(option);
    });
}

// Setup event listeners
function setupEventListeners() {
    searchInput.addEventListener('input', applyFilters);
    shelfFilter.addEventListener('change', applyFilters);
    headerCheckbox.addEventListener('change', toggleAllCheckboxes);
    selectAllBtn.addEventListener('click', () => selectAll(true));
    deselectAllBtn.addEventListener('click', () => selectAll(false));
    printSelectedBtn.addEventListener('click', printLabels);

    // Sortable columns
    document.querySelectorAll('.sortable').forEach(th => {
        th.addEventListener('click', () => {
            const column = th.dataset.column;
            handleSort(column);
        });
    });
}

// Apply filters
function applyFilters() {
    const searchTerm = searchInput.value.toLowerCase().trim();
    const selectedShelf = shelfFilter.value;

    filteredData = allData.filter(item => {
        const matchesSearch = !searchTerm ||
            item.box.toLowerCase().includes(searchTerm) ||
            item.shelf.toLowerCase().includes(searchTerm);

        const matchesShelf = !selectedShelf || item.shelf === selectedShelf;

        return matchesSearch && matchesShelf;
    });

    renderTable();
    updateStats();
}

// Handle sorting
function handleSort(column) {
    if (sortColumn === column) {
        sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        sortColumn = column;
        sortDirection = 'asc';
    }

    filteredData.sort((a, b) => {
        const aVal = a[column].toString().toLowerCase();
        const bVal = b[column].toString().toLowerCase();

        if (aVal < bVal) return sortDirection === 'asc' ? -1 : 1;
        if (aVal > bVal) return sortDirection === 'asc' ? 1 : -1;
        return 0;
    });

    // Update sort icons
    document.querySelectorAll('.sortable').forEach(th => {
        th.classList.remove('active');
        const icon = th.querySelector('.sort-icon');
        icon.textContent = '⇅';
    });

    const activeTh = document.querySelector(`[data-column="${column}"]`);
    activeTh.classList.add('active');
    const icon = activeTh.querySelector('.sort-icon');
    icon.textContent = sortDirection === 'asc' ? '↑' : '↓';

    renderTable();
}

// Render table
function renderTable() {
    if (filteredData.length === 0) {
        tableBody.innerHTML = `
            <tr>
                <td colspan="3" class="no-data">
                    ไม่พบข้อมูล
                </td>
            </tr>
        `;
        return;
    }

    tableBody.innerHTML = filteredData.map(item => `
        <tr class="${item.selected ? 'selected' : ''}" data-id="${item.id}">
            <td>
                <input type="checkbox" 
                       class="row-checkbox" 
                       data-id="${item.id}" 
                       ${item.selected ? 'checked' : ''}>
            </td>
            <td>${escapeHtml(item.box)}</td>
            <td>${escapeHtml(item.shelf)}</td>
        </tr>
    `).join('');

    // Add event listeners to checkboxes
    document.querySelectorAll('.row-checkbox').forEach(checkbox => {
        checkbox.addEventListener('change', handleRowCheckbox);
    });
}

// Handle row checkbox change
function handleRowCheckbox(e) {
    const id = parseInt(e.target.dataset.id);
    const item = allData.find(item => item.id === id);
    if (item) {
        item.selected = e.target.checked;
        const row = e.target.closest('tr');
        row.classList.toggle('selected', e.target.checked);
        updateStats();
        updateHeaderCheckbox();
    }
}

// Toggle all checkboxes
function toggleAllCheckboxes(e) {
    const checked = e.target.checked;

    // First, deselect ALL items (including those not in current filter)
    allData.forEach(item => {
        item.selected = false;
    });

    // Then, select only the filtered items if checked
    if (checked) {
        filteredData.forEach(item => {
            const originalItem = allData.find(i => i.id === item.id);
            if (originalItem) {
                originalItem.selected = true;
            }
        });
    }

    renderTable();
    updateStats();
}

// Select/deselect all
function selectAll(selected) {
    // First, deselect ALL items (including those not in current filter)
    allData.forEach(item => {
        item.selected = false;
    });

    // Then, select/deselect only the filtered items
    if (selected) {
        filteredData.forEach(item => {
            const originalItem = allData.find(i => i.id === item.id);
            if (originalItem) {
                originalItem.selected = true;
            }
        });
    }

    renderTable();
    updateStats();
    updateHeaderCheckbox();
}

// Update header checkbox state
function updateHeaderCheckbox() {
    const visibleItems = filteredData.map(item => allData.find(i => i.id === item.id));
    const allSelected = visibleItems.length > 0 && visibleItems.every(item => item.selected);
    const someSelected = visibleItems.some(item => item.selected);

    headerCheckbox.checked = allSelected;
    headerCheckbox.indeterminate = someSelected && !allSelected;
}

// Update statistics
function updateStats() {
    const selectedCount = allData.filter(item => item.selected).length;

    totalItemsEl.textContent = allData.length;
    selectedItemsEl.textContent = selectedCount;
    filteredItemsEl.textContent = filteredData.length;

    updateHeaderCheckbox();
}

// Print labels for selected items
function printLabels() {
    const selectedItems = allData.filter(item => item.selected);

    console.log('Total items:', allData.length);
    console.log('Selected items:', selectedItems.length);
    console.log('Selected items data:', selectedItems);

    if (selectedItems.length === 0) {
        alert('กรุณาเลือกรายการที่ต้องการพิมพ์');
        return;
    }

    // Create print window
    const printWindow = window.open('', '_blank');

    // Generate HTML for labels
    const labelsHTML = selectedItems.map((item, index) => {
        const qrId = `qr-${index}`;
        const barcodeId = `barcode-${index}`;

        return `
            <div class="label">
                <div class="label-main">
                    <div class="label-text">
                        <h1>${escapeHtml(item.box)}</h1>
                        <div class="codes-row">
                            <div class="barcode-container">
                                <svg id="${barcodeId}" class="barcode"></svg>
                            </div>
                            <div class="qr-container">
                                <div id="${qrId}" class="qr-code"></div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;
    }).join('');

    // Write HTML to print window
    printWindow.document.write(`
        <!DOCTYPE html>
        <html lang="th">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Print Labels</title>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
            <script src="https://cdn.jsdelivr.net/npm/jsbarcode@3.11.5/dist/JsBarcode.all.min.js"></script>
            <style>
                * {
                    margin: 0;
                    padding: 0;
                    box-sizing: border-box;
                }
                
                @page {
                    size: 100mm 80mm;
                    margin: 0;
                }
                
                body {
                    font-family: 'Arial', 'Helvetica', sans-serif;
                    margin: 0;
                    padding: 0;
                }
                
                .label {
                    width: 100mm;
                    height: 80mm;
                    padding: 1mm 8mm 8mm 8mm;
                    page-break-after: always;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    border: 1px dashed #ccc;
                    background: white;
                }
                
                .label:last-child {
                    page-break-after: auto;
                }
                
                .label-main {
                    width: 100%;
                    height: 100%;
                    display: flex;
                    flex-direction: column;
                    align-items: center;
                    justify-content: flex-start;
                    gap: 0mm;
                }
                
                .label-text {
                    width: 100%;
                    display: flex;
                    flex-direction: column;
                    align-items: center;
                    justify-content: flex-start;
                    gap: 0mm;
                }
                
                .label-text h1 {
                    font-size: 52pt;
                    font-weight: bold;
                    letter-spacing: 0.1px;
                    margin: 0;
                    padding: 0 2mm;
                    line-height: 1.1;
                    text-align: center;
                    width: 100%;
                    white-space: nowrap;
                    overflow: visible;
                }
                
                .codes-row {
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    gap: 8mm;
                    margin-top: 0mm;
                    width: 100%;
                }
                
                .barcode-container {
                }
                
                .barcode {
                    width: 100%;
                    max-width: 45mm;
                    height: auto;
                }
                
                .qr-container {
                    display: flex;
                    align-items: center;
                    justify-content: center;
                }
                
                .qr-code {
                    display: inline-block;
                    background: white;
                }
                
                .qr-code canvas {
                    display: block;
                }
                
                @media print {
                    body {
                        margin: 0;
                        padding: 0;
                    }
                    
                    .label {
                        border: none;
                    }
                }
            </style>
        </head>
        <body>
            ${labelsHTML}
            <script>
                // Wait for libraries to load
                window.onload = function() {
                    ${selectedItems.map((item, index) => `
                        // Generate QR Code for label ${index}
                        new QRCode(document.getElementById('qr-${index}'), {
                            text: '${item.box.replace(/'/g, "\\'")}',
                            width: 40,
                            height: 40,
                            colorDark: '#000000',
                            colorLight: '#ffffff',
                            correctLevel: QRCode.CorrectLevel.H
                        });
                        
                        // Generate Barcode for label ${index}
                        JsBarcode('#barcode-${index}', '${item.box.replace(/'/g, "\\'")}', {
                            format: 'CODE128',
                            width: 1.5,
                            height: 40,
                            displayValue: false,
                            margin: 0
                        });
                    `).join('\n')}
                    
                    // Auto print after a short delay to ensure all codes are rendered
                    setTimeout(function() {
                        window.print();
                    }, 500);
                };
            </script>
        </body>
        </html>
    `);

    printWindow.document.close();
}

// Utility function to escape HTML
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Show error message
function showError(message) {
    tableBody.innerHTML = `
        <tr>
            <td colspan="3" class="no-data" style="color: var(--danger-color);">
                ⚠️ ${message}
            </td>
        </tr>
    `;
}

// Initialize app when DOM is ready
document.addEventListener('DOMContentLoaded', init);
