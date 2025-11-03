const { ipcRenderer } = require('electron');
const { jsPDF } = require('jspdf');
const QRCode = require('qrcode');
const fs = require('fs');
const path = require('path');
const remote = require('@electron/remote');

// Global variables - Initialize at the top
let selectedRow = null;
let items = [];
let editingFileName = null;

// Function to load the Arabic font (Amiri)
function loadArabicFont() {
    try {
        const fontPath = getResourcePath('Amiri-Regular.ttf');
        if (fs.existsSync(fontPath)) {
            console.log('Amiri font file found:', fontPath);
            return true;
        } else {
            console.warn('Amiri font file not found:', fontPath);
            return false;
        }
    } catch (error) {
        console.error('Error loading Arabic font:', error);
        return false;
    }
}

// Function to add Arabic font to jsPDF document
function addArabicFontToDoc(doc) {
    try {
        const fontPath = getResourcePath('Amiri-Regular.ttf');
        if (fs.existsSync(fontPath)) {
            // Register the Amiri font with jsPDF
            doc.addFont(fontPath, 'Amiri', 'normal');
            doc.addFont(fontPath, 'Amiri', 'bold');
            console.log('Amiri font added to jsPDF document');
            return true;
        } else {
            console.warn('Amiri font file not found, using default font');
            return false;
        }
    } catch (error) {
        console.error('Error adding Arabic font to jsPDF:', error);
        return false;
    }
}

// Helper function to render bilingual text properly in jsPDF
function renderBilingualText(doc, text, x, y, options = {}) {
    try {
        // Check if text contains Arabic characters
        const arabicRegex = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;
        const hasArabic = arabicRegex.test(text);
        
        if (!hasArabic) {
            // No Arabic text, render normally with default font
            doc.setFont('helvetica');
            doc.text(text, x, y, options);
            return;
        }
        
        // Save current font
        const currentFont = doc.getFont();
        
        // Split text into English and Arabic parts
        const parts = [];
        let currentPart = '';
        let isArabic = false;
        
        for (let i = 0; i < text.length; i++) {
            const char = text[i];
            const charIsArabic = arabicRegex.test(char);
            
            if (i === 0) {
                isArabic = charIsArabic;
                currentPart = char;
            } else if (charIsArabic === isArabic) {
                currentPart += char;
            } else {
                if (currentPart.trim()) {
                    parts.push({ text: currentPart, isArabic: isArabic });
                }
                currentPart = char;
                isArabic = charIsArabic;
            }
        }
        if (currentPart.trim()) {
            parts.push({ text: currentPart, isArabic: isArabic });
        }
        
        // Handle rendering based on alignment
        const align = options.align || 'left';
        const pageWidth = doc.internal.pageSize.width;
        
        if (align === 'center') {
            // For center alignment, split English and Arabic
            const englishParts = parts.filter(p => !p.isArabic).map(p => p.text);
            const arabicParts = parts.filter(p => p.isArabic).map(p => p.text);
            
            const englishText = englishParts.join('').trim();
            const arabicText = arabicParts.join('').trim();
            
            // Calculate positions
            if (englishText && arabicText) {
                // Both English and Arabic - render English on left, Arabic on right
                // Set English font
                doc.setFont('helvetica');
                const englishWidth = doc.getTextWidth(englishText);
                const centerX = pageWidth / 2;
                const englishX = centerX - englishWidth / 2 - 5;
                doc.text(englishText, englishX, y, { ...options, align: 'left' });
                
                // Set Arabic font and render
                doc.setFont('Amiri');
                // Arabic text is already in correct order for display
                const arabicWidth = doc.getTextWidth(arabicText);
                const arabicX = centerX + arabicWidth / 2 + 5;
                doc.text(arabicText, arabicX, y, { ...options, align: 'right' });
            } else if (englishText) {
                // Only English
                doc.setFont('helvetica');
                doc.text(englishText, x, y, { ...options, align: 'center' });
            } else if (arabicText) {
                // Only Arabic
                doc.setFont('Amiri');
                doc.text(arabicText, x, y, { ...options, align: 'center' });
            }
        } else if (align === 'right') {
            // Right alignment - render Arabic first (RTL), then English
            let currentX = x;
            for (let i = parts.length - 1; i >= 0; i--) {
                const part = parts[i];
                if (part.isArabic) {
                    doc.setFont('Amiri');
                    const width = doc.getTextWidth(part.text);
                    currentX -= width;
                    doc.text(part.text, currentX, y, { ...options, align: 'left' });
                } else {
                    doc.setFont('helvetica');
                    const width = doc.getTextWidth(part.text);
                    currentX -= width;
                    doc.text(part.text, currentX, y, { ...options, align: 'left' });
                }
            }
        } else {
            // Left alignment - render as is
            let currentX = x;
            for (const part of parts) {
                if (part.isArabic) {
                    doc.setFont('Amiri');
                    doc.text(part.text, currentX, y, { ...options, align: 'left' });
                    currentX += doc.getTextWidth(part.text) + 1;
                } else {
                    doc.setFont('helvetica');
                    doc.text(part.text, currentX, y, { ...options, align: 'left' });
                    currentX += doc.getTextWidth(part.text) + 1;
                }
            }
        }
        
        // Restore original font
        doc.setFont(currentFont.fontName, currentFont.fontStyle);
    } catch (error) {
        console.error('Error rendering bilingual text:', error);
        // Fallback to normal rendering
        doc.setFont('helvetica');
        doc.text(text, x, y, options);
    }
}

// Function to get the correct resource path in both dev and prod
function getResourcePath(filename) {
    if (
        process.env.NODE_ENV === 'development' ||
        process.defaultApp ||
        /node_modules[\\\/]electron[\\\/]/.test(process.execPath)
    ) {
        return path.join(__dirname, filename);
    }
    return path.join(process.resourcesPath, filename);
}

// Function to handle file reading with error handling
function readImageFile(filepath) {
    try {
        if (!fs.existsSync(filepath)) {
            console.error(`File not found: ${filepath}`);
            return null;
        }
        return fs.readFileSync(filepath);
    } catch (error) {
        console.error(`Error reading file ${filepath}:`, error);
        return null;
    }
}

// Initialize the invoice form
document.addEventListener('DOMContentLoaded', () => {
    initializeInvoiceForm();
});

function initializeInvoiceForm() {
    setCurrentDateTime();
    enableInputFields();
    setupCalculationListeners();
    checkForEditMode();
    calculateProductTotals();
}

function setCurrentDateTime() {
    const now = new Date();
    const pad = n => n.toString().padStart(2, '0');
    const currentDateTime = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}T${pad(now.getHours())}:${pad(now.getMinutes())}`;
    
    const dateInput = document.getElementById('date');
    if (dateInput) {
        dateInput.value = currentDateTime;
    }
}

function enableInputFields() {
    const inputFields = document.querySelectorAll('input[type="text"], input[type="number"], input[type="date"]');
    const readonlyFields = ['subtotal', 'tax', 'total', 'grandTotal', 'balanceDue', 
                            'vatPercent', 'vatAmount', 'totalAmount', 'totalAmountExclVat'];
    
    inputFields.forEach(input => {
        if (!readonlyFields.includes(input.id)) {
            input.removeAttribute('readonly');
        }
    });
}

function setupCalculationListeners() {
    const quantityInput = document.getElementById('quantity');
    const unitPriceInput = document.getElementById('unitPrice');
    const vatPercentInput = document.getElementById('vatPercent');
    
    if (quantityInput) {
        quantityInput.addEventListener('input', calculateProductTotals);
        quantityInput.addEventListener('change', calculateProductTotals);
    }
    
    if (unitPriceInput) {
        unitPriceInput.addEventListener('input', calculateProductTotals);
        unitPriceInput.addEventListener('change', calculateProductTotals);
    }
    
    if (vatPercentInput) {
        vatPercentInput.addEventListener('input', calculateProductTotals);
        vatPercentInput.addEventListener('change', calculateProductTotals);
    }
}

function calculateProductTotals() {
    const quantityInput = document.getElementById('quantity');
    const unitPriceInput = document.getElementById('unitPrice');
    const vatPercentInput = document.getElementById('vatPercent');
    const vatAmountInput = document.getElementById('vatAmount');
    const totalAmountInput = document.getElementById('totalAmount');
    const totalAmountExclVatInput = document.getElementById('totalAmountExclVat');
    
    if (!quantityInput || !unitPriceInput) {
        console.warn('Required input fields not found');
        return;
    }
    
    const quantity = parseFloat(quantityInput.value) || 0;
    const unitPrice = parseFloat(unitPriceInput.value) || 0;
    const vatPercent = parseFloat(vatPercentInput?.value) || 15;
    
    const totalExclVat = quantity * unitPrice;
    const vatAmount = (totalExclVat * vatPercent) / 100;
    const totalInclVat = totalExclVat + vatAmount;
    
    if (totalAmountExclVatInput) totalAmountExclVatInput.value = formatNumber(totalExclVat);
    if (vatAmountInput) vatAmountInput.value = formatNumber(vatAmount);
    if (totalAmountInput) totalAmountInput.value = formatNumber(totalInclVat);
    
    updateInvoiceSummary();
}

// Alias for backward compatibility
const calculateTotals = calculateProductTotals;

function updateInvoiceSummary() {
    let subtotal = 0;
    
    if (items && items.length > 0) {
        subtotal = items.reduce((sum, item) => sum + (item.quantity * item.unitPrice), 0);
    } else {
        const quantity = parseFloat(document.getElementById('quantity')?.value) || 0;
        const unitPrice = parseFloat(document.getElementById('unitPrice')?.value) || 0;
        subtotal = quantity * unitPrice;
    }
    
    const tax = subtotal * 0.15;
    const total = subtotal + tax;
    
    const subtotalEl = document.getElementById('subtotal');
    const taxEl = document.getElementById('tax');
    const totalEl = document.getElementById('total');
    const grandTotalEl = document.getElementById('grandTotal');
    const balanceDueEl = document.getElementById('balanceDue');
    
    if (subtotalEl) subtotalEl.value = formatCurrency(subtotal);
    if (taxEl) taxEl.value = formatCurrency(tax);
    if (totalEl) totalEl.value = formatCurrency(total);
    if (grandTotalEl) grandTotalEl.value = formatCurrency(total);
    if (balanceDueEl) balanceDueEl.value = formatCurrency(total);
}

// Alias for backward compatibility
const updateSummary = updateInvoiceSummary;

function getItemsFromTable() {
    const items = [];
    const tbody = document.querySelector('#itemsTable tbody');
    
    if (!tbody) return items;
    
    const rows = tbody.getElementsByTagName('tr');
    
    for (let i = 0; i < rows.length; i++) {
        const cells = rows[i].getElementsByTagName('td');
        if (cells.length >= 8) {
            items.push({
                productName: cells[0].textContent.trim(),
                unit: cells[1].textContent.trim(),
                quantity: parseFloat(cells[2].textContent.trim()) || 0,
                unitPrice: parseFloat(cells[3].textContent.trim()) || 0,
                vatPercent: parseFloat(cells[4].textContent.replace('%', '').trim()) || 15,
                vatAmount: parseFloat(cells[5].textContent.trim()) || 0,
                totalAmountExclVat: parseFloat(cells[6].textContent.trim()) || 0,
                totalAmount: parseFloat(cells[7].textContent.trim()) || 0,
                subtotal: parseFloat(cells[6].textContent.trim()) || 0
            });
        }
    }
    
    return items;
}

function formatNumber(number) {
    return Number(number).toFixed(2);
}

function formatCurrency(number) {
    return new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'SAR',
        currencyDisplay: 'code',
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    }).format(number).replace('SAR', 'SAR ');
}

function clearProductInputs() {
    const fields = ['productName', 'unit', 'quantity', 'unitPrice'];
    
    fields.forEach(fieldId => {
        const input = document.getElementById(fieldId);
        if (input) {
            input.value = fieldId === 'quantity' || fieldId === 'unitPrice' ? '0' : '';
        }
    });
    
    selectedRow = null;
    calculateProductTotals();
}

function checkForEditMode() {
    const editInvoiceData = localStorage.getItem('editInvoice');
    if (editInvoiceData) {
        try {
            const invoice = JSON.parse(editInvoiceData);
            loadInvoiceForEdit(invoice);
            if (invoice.fileName) {
                editingFileName = invoice.fileName;
            }
            localStorage.removeItem('editInvoice');
        } catch (error) {
            console.error('Error loading invoice for editing:', error);
        }
    }
}

function loadInvoiceForEdit(invoice) {
    const companyInput = document.getElementById('company');
    const addressInput = document.getElementById('address');
    const vatInput = document.getElementById('vat');
    const crNumberInput = document.getElementById('crNumber');
    const supplyDateInput = document.getElementById('supplyDate');
    const dueDateInput = document.getElementById('dueDate');
    const contractNumberInput = document.getElementById('contractNumber');
    const invoicePeriodInput = document.getElementById('invoicePeriod');
    const projectNumberInput = document.getElementById('projectNumber');
    
    if (companyInput) companyInput.value = invoice.customerName || invoice.company || '';
    if (addressInput) addressInput.value = invoice.address || '';
    if (vatInput) vatInput.value = invoice.vatNumber || '';
    if (crNumberInput) crNumberInput.value = invoice.crNumber || '';
    if (supplyDateInput) supplyDateInput.value = invoice.supplyDate || '';
    if (dueDateInput) dueDateInput.value = invoice.dueDate || '';
    if (contractNumberInput) contractNumberInput.value = invoice.contractNumber || '';
    if (invoicePeriodInput) invoicePeriodInput.value = invoice.invoicePeriod || '';
    if (projectNumberInput) projectNumberInput.value = invoice.projectNumber || '';
    
    if (invoice.items && Array.isArray(invoice.items)) {
        const tbody = document.querySelector('#itemsTable tbody');
        if (tbody) {
            tbody.innerHTML = '';
            
            items = invoice.items.map(item => {
                const productName = item.productName || item.name || '';
                const unit = item.unit || item.productCode || item.code || '';
                const quantity = item.quantity || 0;
                const unitPrice = item.unitPrice || 0;
                const vatPercent = item.vatPercent || 15;
                const totalAmountExclVat = item.totalAmountExclVat || item.subtotal || (quantity * unitPrice);
                const vatAmount = item.vatAmount || (totalAmountExclVat * vatPercent / 100);
                const totalAmount = item.totalAmount || (totalAmountExclVat + vatAmount);
                
                const row = tbody.insertRow();
                row.innerHTML = `
                    <td>${productName}</td>
                    <td>${unit}</td>
                    <td>${formatNumber(quantity)}</td>
                    <td>${formatNumber(unitPrice)}</td>
                    <td>${formatNumber(vatPercent)}%</td>
                    <td>${formatNumber(vatAmount)}</td>
                    <td>${formatNumber(totalAmountExclVat)}</td>
                    <td>${formatNumber(totalAmount)}</td>
                `;
                
                row.addEventListener('click', () => {
                    selectedRow = row;
                    document.getElementById('productName').value = productName;
                    document.getElementById('unit').value = unit;
                    document.getElementById('quantity').value = quantity;
                    document.getElementById('unitPrice').value = unitPrice;
                    calculateProductTotals();
                });
                
                return {
                    productName,
                    unit,
                    quantity,
                    unitPrice,
                    vatPercent,
                    vatAmount,
                    totalAmountExclVat,
                    totalAmount,
                    subtotal: totalAmountExclVat
                };
            });
        }
    }
    
    updateInvoiceSummary();
}

// Simple Modal functions
const simpleModal = document.getElementById('simpleModal');
const simpleModalMessage = document.getElementById('simpleModalMessage');
const simpleModalOK = document.getElementById('simpleModalOK');

function showSimpleModal(message) {
    simpleModalMessage.textContent = message;
    simpleModal.style.display = "flex";
}

function hideSimpleModal() {
    simpleModal.style.display = "none";
}

if (simpleModalOK) {
    simpleModalOK.addEventListener('click', hideSimpleModal);
}

if (simpleModal) {
    simpleModal.addEventListener('click', function(e) {
        if (e.target === this) {
            hideSimpleModal();
        }
    });
}

function showAlert(message) {
    let dialogOverlay = document.getElementById('dialogOverlay');
    if (!dialogOverlay) {
        dialogOverlay = document.createElement('div');
        dialogOverlay.id = 'dialogOverlay';
        dialogOverlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);display:flex;align-items:center;justify-content:center;z-index:10000;';
        document.body.appendChild(dialogOverlay);
    }

    const dialog = document.createElement('div');
    dialog.className = 'custom-dialog';
    dialog.style.cssText = 'background:white;padding:20px;border-radius:8px;max-width:400px;';
    dialog.innerHTML = `
        <div class="dialog-content">
            <p style="margin:0 0 15px 0;">${message}</p>
            <button style="padding:8px 20px;background:#3238a5;color:white;border:none;border-radius:4px;cursor:pointer;" onclick="this.parentElement.parentElement.remove();document.getElementById('dialogOverlay').style.display='none';">OK</button>
        </div>
    `;
    
    dialogOverlay.style.display = 'flex';
    dialogOverlay.innerHTML = '';
    dialogOverlay.appendChild(dialog);

    const okButton = dialog.querySelector('button');
    okButton.focus();
    
    const closeHandler = (e) => {
        if (e.key === 'Enter' || e.key === 'Escape') {
            dialog.remove();
            dialogOverlay.style.display = 'none';
            document.removeEventListener('keydown', closeHandler);
        }
    };
    document.addEventListener('keydown', closeHandler);
}

function ensureInputsEditable() {
    const inputFields = document.querySelectorAll('input[type="text"], input[type="number"], input[type="date"]');
    inputFields.forEach(input => {
        if (input && !['subtotal', 'tax', 'total', 'grandTotal', 'balanceDue', 'vatPercent', 'vatAmount', 'totalAmount', 'totalAmountExclVat'].includes(input.id)) {
            input.removeAttribute('readonly');
            input.style.backgroundColor = '#fff';
            input.style.cursor = 'text';
            input.style.pointerEvents = 'auto';
        }
    });
}

// Add item to table
const addButton = document.getElementById('btnAdd');
if (addButton) {
    addButton.addEventListener('click', () => {
        try {
            const productNameInput = document.getElementById('productName');
            const unitInput = document.getElementById('unit');
            const quantityInput = document.getElementById('quantity');
            const unitPriceInput = document.getElementById('unitPrice');
            const vatPercentInput = document.getElementById('vatPercent');
            const vatAmountInput = document.getElementById('vatAmount');
            const totalAmountExclVatInput = document.getElementById('totalAmountExclVat');
            const totalAmountInput = document.getElementById('totalAmount');
            const itemsTableBody = document.getElementById('itemsTable')?.getElementsByTagName('tbody')[0];

            if (!productNameInput || !unitInput || !quantityInput || !unitPriceInput || !itemsTableBody) {
                showSimpleModal("System error: Form elements missing");
                return;
            }

            const productName = productNameInput.value.trim();
            const unit = unitInput.value.trim();
            const quantity = parseFloat(quantityInput.value);
            const unitPrice = parseFloat(unitPriceInput.value);
            const vatPercent = parseFloat(vatPercentInput?.value || '15') || 15;
            const totalAmountExclVat = parseFloat(totalAmountExclVatInput?.value || '0') || 0;
            const vatAmount = parseFloat(vatAmountInput?.value || '0') || 0;
            const totalAmount = parseFloat(totalAmountInput?.value || '0') || 0;

            if (!productName) {
                showSimpleModal("Please enter an items description");
                return;
            }
            if (!unit) {
                showSimpleModal("Please enter a unit");
                return;
            }
            if (isNaN(quantity) || quantity <= 0) {
                showSimpleModal("Please enter a valid quantity (must be greater than 0)");
                return;
            }
            if (isNaN(unitPrice) || unitPrice <= 0) {
                showSimpleModal("Please enter a valid unit price (must be greater than 0)");
                return;
            }

            const item = {
                productName,
                unit,
                quantity,
                unitPrice,
                vatPercent,
                vatAmount,
                totalAmountExclVat,
                totalAmount,
                subtotal: totalAmountExclVat
            };

            items.push(item);
            
            const row = itemsTableBody.insertRow();
            row.innerHTML = `
                <td>${item.productName}</td>
                <td>${item.unit}</td>
                <td>${formatNumber(item.quantity)}</td>
                <td>${formatNumber(item.unitPrice)}</td>
                <td>${formatNumber(item.vatPercent)}%</td>
                <td>${formatNumber(item.vatAmount)}</td>
                <td>${formatNumber(item.totalAmountExclVat)}</td>
                <td>${formatNumber(item.totalAmount)}</td>
            `;

            row.addEventListener('click', () => {
                selectedRow = row;
                const index = selectedRow.rowIndex - 1;
                const selectedItem = items[index];
                
                productNameInput.value = selectedItem.productName || '';
                unitInput.value = selectedItem.unit || '';
                quantityInput.value = selectedItem.quantity || 0;
                unitPriceInput.value = selectedItem.unitPrice || 0;
                
                calculateProductTotals();
            });

            clearProductInputs();
            updateInvoiceSummary();
        } catch (error) {
            console.error('Error adding item:', error);
            showSimpleModal("An unexpected error occurred: " + error.message);
        }
    });
}

// Update selected item
const updateButton = document.getElementById('btnUpdate');
if (updateButton) {
    updateButton.addEventListener('click', () => {
        if (!selectedRow) {
            showAlert('Please select an item to update');
            ensureInputsEditable();
            return;
        }

        const index = selectedRow.rowIndex - 1;
        const productNameInput = document.getElementById('productName');
        const unitInput = document.getElementById('unit');
        const quantityInput = document.getElementById('quantity');
        const unitPriceInput = document.getElementById('unitPrice');
        const vatPercentInput = document.getElementById('vatPercent');
        const vatAmountInput = document.getElementById('vatAmount');
        const totalAmountExclVatInput = document.getElementById('totalAmountExclVat');
        const totalAmountInput = document.getElementById('totalAmount');

        const item = {
            productName: productNameInput.value,
            unit: unitInput.value,
            quantity: parseFloat(quantityInput.value),
            unitPrice: parseFloat(unitPriceInput.value),
            vatPercent: parseFloat(vatPercentInput?.value || '15') || 15,
            vatAmount: parseFloat(vatAmountInput?.value || '0') || 0,
            totalAmountExclVat: parseFloat(totalAmountExclVatInput?.value || '0') || 0,
            totalAmount: parseFloat(totalAmountInput?.value || '0') || 0,
            subtotal: parseFloat(totalAmountExclVatInput?.value || '0') || 0
        };

        items[index] = item;
        
        selectedRow.innerHTML = `
            <td>${item.productName}</td>
            <td>${item.unit}</td>
            <td>${formatNumber(item.quantity)}</td>
            <td>${formatNumber(item.unitPrice)}</td>
            <td>${formatNumber(item.vatPercent)}%</td>
            <td>${formatNumber(item.vatAmount)}</td>
            <td>${formatNumber(item.totalAmountExclVat)}</td>
            <td>${formatNumber(item.totalAmount)}</td>
        `;

        clearProductInputs();
        updateInvoiceSummary();
        ensureInputsEditable();
    });
}

// Delete selected item
const deleteButton = document.getElementById('btnDelete');
if (deleteButton) {
    deleteButton.addEventListener('click', () => {
        if (!selectedRow) {
            showAlert('Please select an item to delete');
            ensureInputsEditable();
            return;
        }

        const index = selectedRow.rowIndex - 1;
        items.splice(index, 1);
        selectedRow.parentNode.removeChild(selectedRow);
        clearProductInputs();
        updateInvoiceSummary();
        ensureInputsEditable();
    });
}

function clearForm() {
    const inputFields = document.querySelectorAll('input[type="text"], input[type="number"], input[type="date"]');
    
    inputFields.forEach(input => {
        if (input) {
            input.removeAttribute('readonly');
            
            if (input.type === 'number') {
                input.value = '0';
            } else if (input.type === 'date') {
                input.value = '';
            } else {
                input.value = '';
            }
        }
    });

    items = [];
    clearItemsTable();
    selectedRow = null;
    calculateProductTotals();
}

function clearItemsTable() {
    const tbody = document.getElementById('itemsTable')?.getElementsByTagName('tbody')[0];
    if (tbody) {
        tbody.innerHTML = '';
    }
}

// Generate Invoice
const generateButton = document.getElementById('btnGenerate');
if (generateButton) {
    generateButton.addEventListener('click', async () => {
        try {
            const companyName = document.getElementById('company')?.value;
            if (!companyName || items.length === 0) {
                showAlert('Please add company information and at least one item');
                ensureInputsEditable();
                return;
            }

            // Create invoice object with all fields
            const invoice = {
                customerName: companyName,
                company: companyName,
                address: document.getElementById('address')?.value || '',
                vatNumber: document.getElementById('vat')?.value || '',
                crNumber: document.getElementById('crNumber')?.value || '',
                supplyDate: document.getElementById('supplyDate')?.value || '',
                dueDate: document.getElementById('dueDate')?.value || '',
                contractNumber: document.getElementById('contractNumber')?.value || '',
                invoicePeriod: document.getElementById('invoicePeriod')?.value || '',
                projectNumber: document.getElementById('projectNumber')?.value || '',
                date: new Date().toISOString(),
                items: items,
                subtotal: parseFloat(document.getElementById('subtotal')?.value.replace(/[^0-9.-]+/g, '') || '0'),
                tax: parseFloat(document.getElementById('tax')?.value.replace(/[^0-9.-]+/g, '') || '0'),
                total: parseFloat(document.getElementById('total')?.value.replace(/[^0-9.-]+/g, '') || '0')
            };

            if (editingFileName) {
                invoice.fileName = editingFileName;
            }

            // Generate random invoice number (4 digits)
            const invoiceNumber = String(Math.floor(1000 + Math.random() * 9000));

            // Generate QR code with minimal essential data
            const qrData = `Invoice: ${invoiceNumber}
Date: ${new Date(invoice.date).toLocaleDateString()}
Customer: ${invoice.customerName}
Total: ${formatNumber(invoice.total)}`;

            const qrCanvas = await QRCode.toCanvas(qrData, {
                errorCorrectionLevel: 'M',
                margin: 1,
                width: 150
            });
            const qrImage = qrCanvas.toDataURL('image/png');

            // Generate PDF
            const doc = new jsPDF({
                orientation: 'portrait',
                unit: 'mm',
                format: 'a4'
            });

            try {
                // Load Arabic font
                loadArabicFont();
                
                // Add Arabic font to document
                const hasArabicFont = addArabicFontToDoc(doc);
                if (!hasArabicFont) {
                    console.warn('Arabic font not loaded. Arabic text may not display correctly.');
                }
                
                // Read and add the logo
                const logoPath = getResourcePath('logo.png');
                const logoData = fs.readFileSync(logoPath);
                const logoBase64 = Buffer.from(logoData).toString('base64');
                
                // Add company logo and header
                const fullPageWidth = doc.internal.pageSize.width;
                const logoHeight = 50;
                
                doc.addImage('data:image/png;base64,' + logoBase64, 'PNG', 0, 10, fullPageWidth, logoHeight);

                // Reset text settings
                doc.setTextColor(0, 0, 0);
                doc.setFont('helvetica', 'normal');
                doc.setFontSize(11);

                const headerBottomY = 60;
                let currentY = headerBottomY + 5;

                // Blue header bar for "VAT INVOICE / فاتورة ضريبية"
                doc.setFillColor(70, 130, 180); // Steel blue color
                doc.rect(15, currentY, doc.internal.pageSize.width - 30, 10, 'F');
                
                // White text on blue background
                doc.setTextColor(255, 255, 255);
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(14);
                renderBilingualText(doc, 'VAT INVOICE / فاتورة ضريبية', doc.internal.pageSize.width / 2, currentY + 7, { align: 'center' });
                
                currentY += 12;

                // Reset text color to black
                doc.setTextColor(0, 0, 0);

                // Two-column layout for addresses
                const leftBoxX = 15;
                const rightBoxX = doc.internal.pageSize.width / 2 + 2;
                const boxWidth = (doc.internal.pageSize.width - 34) / 2;
                const boxHeight = 35;

                // Draw boxes for both addresses
                doc.setDrawColor(0, 0, 0);
                doc.setLineWidth(0.3);
                doc.rect(leftBoxX, currentY, boxWidth, boxHeight); // Left box (From)
                doc.rect(rightBoxX, currentY, boxWidth, boxHeight); // Right box (To)

                // LEFT BOX - "From" (Hardcoded company address)
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(9);
                let leftY = currentY + 5;
                
                // From label (English left, Arabic right)
                doc.text('From', leftBoxX + 3, leftY);
                renderBilingualText(doc, 'من عنوان', leftBoxX + boxWidth - 3, leftY, { align: 'right' });
                leftY += 5;

                // Company name
                doc.setFont('helvetica', 'normal');
                doc.setFontSize(8);
                doc.text('JUDE MOHAMMED HUSSEIN AL-SHARIF GEN. CONT. EST.', leftBoxX + 3, leftY, { maxWidth: boxWidth - 6 });
                leftY += 4;

                // Address label
                doc.setFont('helvetica', 'bold');
                doc.text('Address:', leftBoxX + 3, leftY);
                renderBilingualText(doc, 'العنوان', leftBoxX + boxWidth - 3, leftY, { align: 'right' });
                leftY += 4;

                // Address details
                doc.setFont('helvetica', 'normal');
                doc.text('Build No. 2626 - Al Quds, 7847 Abu Shajarah Dist.', leftBoxX + 3, leftY, { maxWidth: boxWidth - 6 });
                leftY += 4;
                doc.text('48321 Umluj - Kingdom Of Saudi Arabia', leftBoxX + 3, leftY, { maxWidth: boxWidth - 6 });
                leftY += 5;

                // VAT Number
                doc.setFont('helvetica', 'bold');
                doc.text('VAT:', leftBoxX + 3, leftY);
                doc.setFont('helvetica', 'normal');
                doc.text('311537435500003', leftBoxX + 12, leftY);
                renderBilingualText(doc, 'الرقم الضريبي', leftBoxX + boxWidth - 3, leftY, { align: 'right' });
                leftY += 4;

                // CR Number
                doc.setFont('helvetica', 'bold');
                doc.text('CR No:', leftBoxX + 3, leftY);
                doc.setFont('helvetica', 'normal');
                doc.text('4701103471', leftBoxX + 15, leftY);
                renderBilingualText(doc, 'السجل التجاري', leftBoxX + boxWidth - 3, leftY, { align: 'right' });

                // RIGHT BOX - "To" (Customer address from form)
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(9);
                let rightY = currentY + 5;
                
                // To label (English left, Arabic right)
                doc.text('To', rightBoxX + 3, rightY);
                renderBilingualText(doc, 'إلى', rightBoxX + boxWidth - 3, rightY, { align: 'right' });
                rightY += 5;

                // Customer/Company name
                doc.setFont('helvetica', 'normal');
                doc.setFontSize(8);
                doc.text(invoice.customerName, rightBoxX + 3, rightY, { maxWidth: boxWidth - 6 });
                rightY += 5;

                // Address label (if address exists)
                if (invoice.address) {
                    doc.setFont('helvetica', 'bold');
                    doc.text('Address:', rightBoxX + 3, rightY);
                    renderBilingualText(doc, 'العنوان', rightBoxX + boxWidth - 3, rightY, { align: 'right' });
                    rightY += 4;

                    // Address details
                    doc.setFont('helvetica', 'normal');
                    doc.text(invoice.address, rightBoxX + 3, rightY, { maxWidth: boxWidth - 6 });
                    rightY += 5;
                }

                // VAT Number
                if (invoice.vatNumber) {
                    doc.setFont('helvetica', 'bold');
                    doc.text('VAT:', rightBoxX + 3, rightY);
                    doc.setFont('helvetica', 'normal');
                    doc.text(invoice.vatNumber, rightBoxX + 12, rightY);
                    renderBilingualText(doc, 'الرقم الضريبي', rightBoxX + boxWidth - 3, rightY, { align: 'right' });
                    rightY += 4;
                }

                // CR Number
                if (invoice.crNumber) {
                    doc.setFont('helvetica', 'bold');
                    doc.text('CR No:', rightBoxX + 3, rightY);
                    doc.setFont('helvetica', 'normal');
                    doc.text(invoice.crNumber, rightBoxX + 15, rightY);
                    renderBilingualText(doc, 'السجل التجاري', rightBoxX + boxWidth - 3, rightY, { align: 'right' });
                }

                currentY += boxHeight + 5;

                // Additional invoice details (after the address boxes)
                // Invoice number and date in header row
                const infoRowY = currentY;
                
                // Left side - Invoice Number
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(9);
                doc.text('Invoice No:', leftBoxX + 3, infoRowY);
                renderBilingualText(doc, 'رقم الفاتورة', leftBoxX + boxWidth - 3, infoRowY, { align: 'right' });
                
                doc.setFont('helvetica', 'normal');
                doc.text(invoiceNumber, leftBoxX + 25, infoRowY);
                
                // Right side - Date
                const formattedDate = new Date(invoice.date).toLocaleDateString('en-GB');
                doc.setFont('helvetica', 'bold');
                doc.text('Date:', rightBoxX + 3, infoRowY);
                renderBilingualText(doc, 'التاريخ', rightBoxX + boxWidth - 3, infoRowY, { align: 'right' });
                
                doc.setFont('helvetica', 'normal');
                doc.text(formattedDate, rightBoxX + 20, infoRowY);
                
                currentY += 7;

                // Additional invoice details in rows
                const leftColX = 15;
                const labelWidth = 35;
                const valueStartX = leftColX + labelWidth;
                const rowSpacing = 6;
                let detailY = currentY;

                doc.setFontSize(8);

                // Supply Date
                if (invoice.supplyDate) {
                    doc.setFont('helvetica', 'bold');
                    doc.text('Supply Date:', leftColX, detailY);
                    doc.setFont('helvetica', 'normal');
                    doc.text(new Date(invoice.supplyDate).toLocaleDateString(), valueStartX, detailY);
                    detailY += rowSpacing;
                }

                // Due Date
                if (invoice.dueDate) {
                    doc.setFont('helvetica', 'bold');
                    doc.text('Due Date:', leftColX, detailY);
                    doc.setFont('helvetica', 'normal');
                    doc.text(new Date(invoice.dueDate).toLocaleDateString(), valueStartX, detailY);
                    detailY += rowSpacing;
                }

                // Contract Number
                if (invoice.contractNumber) {
                    doc.setFont('helvetica', 'bold');
                    doc.text('Contract/PO:', leftColX, detailY);
                    doc.setFont('helvetica', 'normal');
                    doc.text(invoice.contractNumber, valueStartX, detailY);
                    detailY += rowSpacing;
                }

                // Invoice Period
                if (invoice.invoicePeriod) {
                    doc.setFont('helvetica', 'bold');
                    doc.text('Invoice Period:', leftColX, detailY);
                    doc.setFont('helvetica', 'normal');
                    doc.text(invoice.invoicePeriod, valueStartX, detailY);
                    detailY += rowSpacing;
                }

                // Project Number
                if (invoice.projectNumber) {
                    doc.setFont('helvetica', 'bold');
                    doc.text('Project/Ref No:', leftColX, detailY);
                    doc.setFont('helvetica', 'normal');
                    doc.text(invoice.projectNumber, valueStartX, detailY);
                    detailY += rowSpacing;
                }

                // Bank details
                detailY += 2;
                doc.setFont('helvetica', 'bold');
                doc.text('Bank Name:', leftColX, detailY);
                doc.setFont('helvetica', 'normal');
                doc.text('Alrajhi Bank', valueStartX, detailY);
                detailY += rowSpacing;

                doc.setFont('helvetica', 'bold');
                doc.text('Account Title:', leftColX, detailY);
                doc.setFont('helvetica', 'normal');
                doc.text('Rakan Hussein Al-Fatih Contracting Company', valueStartX, detailY);
                detailY += rowSpacing;

                doc.setFont('helvetica', 'bold');
                doc.text('IBAN:', leftColX, detailY);
                doc.setFont('helvetica', 'normal');
                doc.text('SA6280000146608016555919', valueStartX, detailY);
                
                // QR code on the right side
                const qrWidth = 35;
                const qrX = doc.internal.pageSize.width - qrWidth - 15;
                const qrY = currentY;
                doc.addImage(qrImage, 'PNG', qrX, qrY, qrWidth, qrWidth);

                detailY += rowSpacing + 8;

                // Items table
                let startY = detailY;
                let startX = 15;
                const tableWidth = doc.internal.pageSize.width - 30;

                // Column widths for 8 columns
                const colWidths = [
                    tableWidth * 0.20, // Items Description
                    tableWidth * 0.10, // Unit
                    tableWidth * 0.10, // Quantity
                    tableWidth * 0.12, // Unit Price
                    tableWidth * 0.08, // VAT %
                    tableWidth * 0.12, // VAT Amount
                    tableWidth * 0.14, // Total excl VAT
                    tableWidth * 0.14  // Total Amount
                ];

                // Table headers
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(9);
                let colX = startX;
                const headers = ['Items Description', 'Unit', 'Quantity/No', 'Unit Price', 'VAT %', 'VAT Amount', 'Total excl VAT', 'Total Amount'];
                headers.forEach((header, i) => {
                    doc.text(header, colX + 2, startY + 6, { maxWidth: colWidths[i] - 4 });
                    colX += colWidths[i];
                });

                // Table content
                doc.setFont('helvetica', 'normal');
                doc.setFontSize(9);
                currentY = startY + 8;
                const pageHeight = doc.internal.pageSize.height;
                const footerHeight = 60;
                const maxContentHeight = pageHeight - footerHeight;

                items.forEach((item, index) => {
                    const wrapText = (text, width) => {
                        let lines = doc.splitTextToSize(text, width - 4);
                        if (lines.length > 2) {
                            lines = [lines[0], lines[1].slice(0, width/2) + '...'];
                        }
                        return lines;
                    };
                    
                    const nameLines = wrapText(item.productName || '', colWidths[0]);
                    const unitLines = wrapText(item.unit || '', colWidths[1]);
                    const rowHeight = Math.max(8, nameLines.length * 7, unitLines.length * 7);

                    // Check if new page needed
                    if (currentY + rowHeight > maxContentHeight) {
                        doc.addPage();
                        currentY = 30;
                        
                        // Redraw headers
                        doc.setFont('helvetica', 'bold');
                        doc.setFontSize(9);
                        colX = startX;
                        headers.forEach((header, i) => {
                            doc.text(header, colX + 2, currentY + 6, { maxWidth: colWidths[i] - 4 });
                            colX += colWidths[i];
                        });
                        
                        currentY += 8;
                        doc.setFont('helvetica', 'normal');
                    }

                    // Draw cell borders
                    colX = startX;
                    for (let i = 0; i < colWidths.length; i++) {
                        doc.rect(colX, currentY, colWidths[i], rowHeight);
                        colX += colWidths[i];
                    }

                    // Draw content
                    const vCenter = currentY + rowHeight / 2 + 2;
                    
                    // Items Description (wrapped)
                    nameLines.forEach((line, idx) => {
                        doc.text(line, startX + 2, currentY + 6 + (idx * 7), { maxWidth: colWidths[0] - 4 });
                    });
                    
                    // Unit (wrapped)
                    unitLines.forEach((line, idx) => {
                        doc.text(line, startX + colWidths[0] + 2, currentY + 6 + (idx * 7), { maxWidth: colWidths[1] - 4 });
                    });
                    
                    // Other columns (centered vertically)
                    doc.text(formatNumber(item.quantity), startX + colWidths[0] + colWidths[1] + 2, vCenter);
                    doc.text(formatNumber(item.unitPrice), startX + colWidths[0] + colWidths[1] + colWidths[2] + 2, vCenter);
                    doc.text(formatNumber(item.vatPercent) + '%', startX + colWidths[0] + colWidths[1] + colWidths[2] + colWidths[3] + 2, vCenter);
                    doc.text(formatNumber(item.vatAmount), startX + colWidths[0] + colWidths[1] + colWidths[2] + colWidths[3] + colWidths[4] + 2, vCenter);
                    doc.text(formatNumber(item.totalAmountExclVat), startX + colWidths[0] + colWidths[1] + colWidths[2] + colWidths[3] + colWidths[4] + colWidths[5] + 2, vCenter);
                    doc.text(formatNumber(item.totalAmount), startX + colWidths[0] + colWidths[1] + colWidths[2] + colWidths[3] + colWidths[4] + colWidths[5] + colWidths[6] + 2, vCenter);

                    currentY += rowHeight;
                });

                // Summary section
                const summaryStartY = currentY + 10;
                const summaryWidth = 80;
                const summaryX = doc.internal.pageSize.width - summaryWidth - 15;

                // Check if summary fits
                if (summaryStartY + 40 > maxContentHeight) {
                    doc.addPage();
                    currentY = 30;
                }

                // Draw summary box
                doc.rect(summaryX, currentY, summaryWidth, 30);

                // Summary labels
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(10);
                doc.text('Subtotal:', summaryX + 5, currentY + 8);
                doc.text('VAT (15%):', summaryX + 5, currentY + 16);
                doc.text('Total Amount:', summaryX + 5, currentY + 24);

                // Summary values
                doc.setFont('helvetica', 'normal');
                doc.text(formatCurrency(invoice.subtotal), summaryX + summaryWidth - 5, currentY + 8, { align: 'right' });
                doc.text(formatCurrency(invoice.tax), summaryX + summaryWidth - 5, currentY + 16, { align: 'right' });
                doc.text(formatCurrency(invoice.total), summaryX + summaryWidth - 5, currentY + 24, { align: 'right' });

                // Total in words
                doc.setFont('helvetica', 'italic');
                doc.setFontSize(9);
                const totalInWords = `(${numberToWords(Math.round(invoice.total * 100) / 100)})`;
                doc.text(totalInWords, 15, currentY + 40);

                // Add page numbers
                const totalPages = doc.internal.getNumberOfPages();
                for (let i = 1; i <= totalPages; i++) {
                    doc.setPage(i);
                    doc.setFont('helvetica', 'normal');
                    doc.setFontSize(10);
                    const pageText = `Page ${i} of ${totalPages}`;
                    const textWidth = doc.getStringUnitWidth(pageText) * doc.internal.getFontSize() / doc.internal.scaleFactor;
                    const textX = (doc.internal.pageSize.width - textWidth) / 2;
                    doc.text(pageText, textX, doc.internal.pageSize.height - 35);
                }

                // Add footer image on last page
                try {
                    const footerPath = getResourcePath('footer.png');
                    const footerData = fs.readFileSync(footerPath);
                    const footerBase64 = Buffer.from(footerData).toString('base64');
                    const footerY = doc.internal.pageSize.height - 30;
                    doc.addImage('data:image/png;base64,' + footerBase64, 'PNG', 0, footerY, doc.internal.pageSize.width, 30);
                } catch (error) {
                    console.error('Error adding footer:', error);
                }

                // Save invoice data
                await ipcRenderer.invoke('save-invoice', invoice);

                // Save PDF
                const pdfBuffer = doc.output('arraybuffer');
                const pdfPath = await ipcRenderer.invoke('save-pdf', pdfBuffer);

                showAlert('Invoice has been generated successfully!');
                ensureInputsEditable();

                // Clear the form after successful generation
                clearForm();
                clearItemsTable();
                editingFileName = null;

            } catch (error) {
                console.error('Error generating invoice:', error);
                showAlert('Error generating invoice: ' + (error && error.message ? error.message : error));
                ensureInputsEditable();
            }
        } catch (error) {
            console.error('Error generating invoice:', error);
            showAlert('Error generating invoice: ' + error.message);
            ensureInputsEditable();
        }
    });
}

// Helper function to convert number to words
function numberToWords(number) {
    const ones = ['', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine'];
    const tens = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];
    const teens = ['ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen'];
    
    if (number === 0) return 'zero';
    
    const wholePart = Math.floor(number);
    const decimalPart = Math.round((number - wholePart) * 100);
    
    let result = wholePart.toString();
    
    if (decimalPart > 0) {
        result += ' and ' + decimalPart + ' cents';
    }
    
    return result;
}