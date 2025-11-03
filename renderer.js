const { ipcRenderer } = require('electron');
const { jsPDF } = require('jspdf');
const QRCode = require('qrcode');
const fs = require('fs');
const path = require('path');
const remote = require('@electron/remote');

// Function to get the correct resource path in both dev and prod
function getResourcePath(filename) {
    // If running in development or unpackaged mode
    if (
        process.env.NODE_ENV === 'development' ||
        process.defaultApp ||
        /node_modules[\\\/]electron[\\\/]/.test(process.execPath)
    ) {
        return path.join(__dirname, filename);
    }
    // If running as a packaged app
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

// Initialize date field with current date and setup input fields
document.addEventListener('DOMContentLoaded', () => {
    // Set initial date with current date and time in datetime-local format
    const now = new Date();
    const pad = n => n.toString().padStart(2, '0');
    const currentDateTime = `${now.getFullYear()}-${pad(now.getMonth()+1)}-${pad(now.getDate())}T${pad(now.getHours())}:${pad(now.getMinutes())}`;
    document.getElementById('date').value = currentDateTime;
    
    // Enable all input fields
    const inputFields = document.querySelectorAll('input[type="text"], input[type="number"]');
    inputFields.forEach(input => {
        // Only remove readonly from fields that are NOT subtotal, tax, or total
        if (input && !['subtotal', 'tax', 'total'].includes(input.id)) {
            input.removeAttribute('readonly');
            // Clear any existing event listeners
            input.replaceWith(input.cloneNode(true));
            // Get the fresh input reference
            const freshInput = document.getElementById(input.id);
            if (freshInput) {
                // Add new event listeners
                freshInput.addEventListener('focus', (e) => {
                    e.target.removeAttribute('readonly');
                });
                freshInput.addEventListener('blur', (e) => {
                    if (!e.target.hasAttribute('readonly')) {
                        e.target.removeAttribute('readonly');
                    }
                });
            }
        }
    });

    // Check for invoice to edit
    const editInvoiceData = localStorage.getItem('editInvoice');
    if (editInvoiceData) {
        try {
            const invoice = JSON.parse(editInvoiceData);
            loadInvoiceForEdit(invoice);
            // Store the fileName for update
            if (invoice.fileName) {
                editingFileName = invoice.fileName;
            }
            localStorage.removeItem('editInvoice');
        } catch (error) {
            console.error('Error loading invoice for editing:', error);
        }
    }
});

// Function to load invoice for editing
function loadInvoiceForEdit(invoice) {
    document.getElementById('customer').value = invoice.customerName;
    document.getElementById('mobile').value = invoice.mobile;
    document.getElementById('vat').value = invoice.vatNumber;
    document.getElementById('date').value = invoice.date;

    items = Array.isArray(invoice.items) ? invoice.items : [];
    const tbody = document.getElementById('itemsTable').getElementsByTagName('tbody')[0];
    tbody.innerHTML = '';
    
    items.forEach(item => {
        const row = tbody.insertRow();
        row.innerHTML = `
            <td>${item.productName || item.name || ''}</td>
            <td>${item.productCode || item.code || ''}</td>
            <td>${formatNumber(item.quantity)} m</td>
            <td>${formatNumber(item.unitPrice)}</td>
            <td>${formatNumber(item.subtotal)}</td>
        `;

        row.addEventListener('click', () => {
            selectedRow = row;
            const index = selectedRow.rowIndex - 1;
            const selectedItem = items[index];
            document.getElementById('productName').value = selectedItem.productName || selectedItem.name || '';
            document.getElementById('productCode').value = selectedItem.productCode || selectedItem.code || '';
            document.getElementById('quantity').value = selectedItem.quantity;
            document.getElementById('unitPrice').value = selectedItem.unitPrice;
        });
    });

    updateSummary();
    ensureInputsEditable();
}

// Global variables
let selectedRow = null;
let items = [];
let editingFileName = null;

// Helper function to format currency
const formatCurrency = (number) => {
    return new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'SAR',
        currencyDisplay: 'code',
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    }).format(number).replace('SAR', 'SAR ');
};

// Helper function to format number with 2 decimal places
const formatNumber = (number) => {
    return Number(number).toFixed(2);
};

// Update summary calculations
const updateSummary = () => {
    const subtotal = items.reduce((sum, item) => sum + (item.quantity * item.unitPrice), 0);
    const tax = subtotal * 0.15;
    const total = subtotal + tax;

    document.getElementById('subtotal').value = formatCurrency(subtotal);
    document.getElementById('tax').value = formatCurrency(tax);
    document.getElementById('total').value = formatCurrency(total);
};

// Clear product input fields
const clearProductInputs = () => {
    document.getElementById('productName').value = '';
    document.getElementById('productCode').value = '';
    document.getElementById('quantity').value = '0';
    document.getElementById('unitPrice').value = '0';
    selectedRow = null;
};

// Add item to table
const addButton = document.getElementById('btnAdd');
const simpleModal = document.getElementById('simpleModal');
const simpleModalMessage = document.getElementById('simpleModalMessage');
const simpleModalOK = document.getElementById('simpleModalOK');

if (addButton) {
    addButton.addEventListener('click', () => {
        try {
            const productNameInput = document.getElementById('productName');
            const productCodeInput = document.getElementById('productCode');
            const quantityInput = document.getElementById('quantity');
            const unitPriceInput = document.getElementById('unitPrice');
            const itemsTableBody = document.getElementById('itemsTable').getElementsByTagName('tbody')[0];

            if (!productNameInput || !productCodeInput || !quantityInput || !unitPriceInput || !itemsTableBody) {
                showSimpleModal("System error: Form elements missing");
                return;
            }

            const productName = productNameInput.value.trim();
            const productCode = productCodeInput.value.trim();
            const quantity = parseFloat(quantityInput.value);
            const unitPrice = parseFloat(unitPriceInput.value);

            // Validate inputs with specific error messages
            if (!productName) {
                showSimpleModal("Please enter a product name");
                return;
            }
            if (!productCode) {
                showSimpleModal("Please enter a product code");
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

            // Rest of your existing code...
            const item = {
                productName,
                productCode,
                quantity,
                unitPrice,
                subtotal: quantity * unitPrice
            };

            items.push(item);
            
            const row = itemsTableBody.insertRow();
            
            row.innerHTML = `
                <td>${item.productName}</td>
                <td>${item.productCode}</td>
                <td>${formatNumber(item.quantity)} m</td>
                <td>${formatNumber(item.unitPrice)}</td>
                <td>${formatNumber(item.subtotal)}</td>
            `;

            row.addEventListener('click', () => {
                selectedRow = row;
                const index = selectedRow.rowIndex - 1;
                const item = items[index];
                
                productNameInput.value = item.productName;
                productCodeInput.value = item.productCode;
                quantityInput.value = item.quantity;
                unitPriceInput.value = item.unitPrice;
            });

            clearProductInputs();
            updateSummary();
        } catch (error) {
            console.error('Error adding item:', error);
            showSimpleModal("An unexpected error occurred");
        }
    });
}

// Simple Modal functions
function showSimpleModal(message) {
    simpleModalMessage.textContent = message;
    simpleModal.style.display = "flex";
}

function hideSimpleModal() {
    simpleModal.style.display = "none";
}

// Setup OK button
simpleModalOK.addEventListener('click', hideSimpleModal);

// Close when clicking outside modal
simpleModal.addEventListener('click', function(e) {
    if (e.target === this) {
        hideSimpleModal();
    }
});

// Custom alert function to replace window.alert
function showAlert(message) {
    // Create a dialog div if it doesn't exist
    let dialogOverlay = document.getElementById('dialogOverlay');
    if (!dialogOverlay) {
        dialogOverlay = document.createElement('div');
        dialogOverlay.id = 'dialogOverlay';
        document.body.appendChild(dialogOverlay);
    }

    const dialog = document.createElement('div');
    dialog.className = 'custom-dialog';
    dialog.innerHTML = `
        <div class="dialog-content">
            <p>${message}</p>
            <button onclick="this.parentElement.parentElement.remove();document.getElementById('dialogOverlay').style.display='none';">OK</button>
        </div>
    `;
    
    dialogOverlay.style.display = 'block';
    dialogOverlay.appendChild(dialog);

    // Focus handling
    const okButton = dialog.querySelector('button');
    okButton.focus();
    
    // Remove dialog on Enter or Escape
    document.addEventListener('keydown', function closeOnKey(e) {
        if (e.key === 'Enter' || e.key === 'Escape') {
            dialog.remove();
            dialogOverlay.style.display = 'none';
            document.removeEventListener('keydown', closeOnKey);
        }
    });
}

// Function to ensure all input fields are editable
function ensureInputsEditable() {
    const inputFields = document.querySelectorAll('input[type="text"], input[type="number"]');
    inputFields.forEach(input => {
        if (input && !['subtotal', 'tax', 'total'].includes(input.id)) {
            input.removeAttribute('readonly');
            input.style.backgroundColor = '#fff';
            input.style.cursor = 'text';
            input.style.pointerEvents = 'auto';
        }
    });
}

// Update selected item
document.getElementById('btnUpdate').addEventListener('click', () => {
    if (!selectedRow) {
        showAlert('Please select an item to update');
        ensureInputsEditable();
        return;
    }

    const index = selectedRow.rowIndex - 1;
    const item = {
        productName: document.getElementById('productName').value,
        productCode: document.getElementById('productCode').value,
        quantity: parseFloat(document.getElementById('quantity').value),
        unitPrice: parseFloat(document.getElementById('unitPrice').value),
        subtotal: parseFloat(document.getElementById('quantity').value) * parseFloat(document.getElementById('unitPrice').value)
    };

    items[index] = item;
    
    selectedRow.innerHTML = `
        <td>${item.productName}</td>
        <td>${item.productCode}</td>
        <td>${formatNumber(item.quantity)} m</td>
        <td>${formatNumber(item.unitPrice)}</td>
        <td>${formatNumber(item.subtotal)}</td>
    `;

    clearProductInputs();
    updateSummary();
    ensureInputsEditable();
});

// Delete selected item
document.getElementById('btnDelete').addEventListener('click', () => {
    if (!selectedRow) {
        showAlert('Please select an item to delete');
        ensureInputsEditable();
        return;
    }

    const index = selectedRow.rowIndex - 1;
    items.splice(index, 1);
    selectedRow.parentNode.removeChild(selectedRow);
    clearProductInputs();
    updateSummary();
    ensureInputsEditable();
});

// Function to clear form and reset fields
function clearForm() {
    // Get all input fields
    const inputFields = document.querySelectorAll('input[type="text"], input[type="number"]');
    
    // Clear and reset each input field
    inputFields.forEach(input => {
        if (input) {
            // Remove readonly attribute
            input.removeAttribute('readonly');
            
            // Reset value based on input type
            if (input.type === 'number') {
                input.value = '0';
            } else if (input.id === 'date') {
                input.valueAsDate = new Date();
            } else {
                input.value = '';
            }
        }
    });

    // Clear items array and table
    items = [];
    clearItemsTable();

    // Reset selected row
    selectedRow = null;
}

// Function to clear items table
function clearItemsTable() {
    const tbody = document.getElementById('itemsTable').getElementsByTagName('tbody')[0];
    if (tbody) {
        tbody.innerHTML = '';
    }
}

// Generate Invoice
document.getElementById('btnGenerate').addEventListener('click', async () => {
    try {
        const customerName = document.getElementById('customer').value;
        if (!customerName || items.length === 0) {
            showAlert('Please add customer information and at least one item');
            ensureInputsEditable();
            return;
        }

        // Create invoice object
        const invoice = {
            customerName,
            mobile: document.getElementById('mobile').value,
            vatNumber: document.getElementById('vat').value,
            date: document.getElementById('date').value,
            items,
            subtotal: parseFloat(document.getElementById('subtotal').value.replace(/[^0-9.-]+/g, '')),
            tax: parseFloat(document.getElementById('tax').value.replace(/[^0-9.-]+/g, '')),
            total: parseFloat(document.getElementById('total').value.replace(/[^0-9.-]+/g, ''))
        };
        if (editingFileName) {
            invoice.fileName = editingFileName;
        }

        // Generate random invoice number (4 digits)
        const invoiceNumber = String(Math.floor(1000 + Math.random() * 9000));

        // Generate QR code with minimal essential data only
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
            // Read and add the logo
            const logoPath = getResourcePath('logo.png');
            const logoData = fs.readFileSync(logoPath);
            const logoBase64 = Buffer.from(logoData).toString('base64');
            
            // Add company logo and header
            const pageCenter = doc.internal.pageSize.width / 2;
            const logoHeight = 50; // Increased height for better visibility
            const fullPageWidth = doc.internal.pageSize.width;
            const pageWidth = fullPageWidth - 30;
            
            // Add logo with full width (no margins)
            doc.addImage('data:image/png;base64,' + logoBase64, 'PNG', 0, 10, fullPageWidth, logoHeight);

            // Reset text color and font for the rest of the invoice
            doc.setTextColor(0, 0, 0);
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(11);

            // After drawing the logo and header lines...
            const headerBottomY = 50; // move everything further down
            const contentStartY = headerBottomY + 12; // more space below logo/header

            // Title and Date
            doc.setFontSize(26);
            doc.setFont('helvetica', 'bold');
            doc.text('Invoice', doc.internal.pageSize.width / 2, contentStartY, { align: 'center' });

            // Add date near the title with bold "Date:" label
            doc.setFontSize(10);
            const formattedDate = new Date(invoice.date).toLocaleString('en-GB', {
                hour: '2-digit',
                minute: '2-digit',
                year: 'numeric',
                month: '2-digit',
                day: '2-digit'
            }).replace(',', '');

            // Make "Date:" bold
            doc.setFont('helvetica', 'bold');
            doc.text('Date:', doc.internal.pageSize.width - 60, contentStartY);

            // Make the actual date value normal weight
            doc.setFont('helvetica', 'normal');
            doc.text(formattedDate, doc.internal.pageSize.width - 60 + 15, contentStartY);

            // Invoice details and QR code
            const detailsStartY = contentStartY + 8;
            const leftColX = 15;
            const labelWidth = 35; // Fixed width for all labels to ensure consistent alignment
            const valueStartX = leftColX + labelWidth; // All values start from this position
            const rightColX = doc.internal.pageSize.width - 60; // QR code right margin
            const rowSpacing = 7;
            let detailY = detailsStartY;

            // QR code (right side)
            const qrWidth = 38;
            const qrY = detailsStartY;
            doc.addImage(qrImage, 'PNG', rightColX, qrY, qrWidth, qrWidth);

            // Details (left stack) - All values aligned at same position
            doc.setFontSize(10);

            // Invoice Number
            doc.setFont('helvetica', 'bold');
            doc.text('Invoice No:', leftColX, detailY);
            doc.setFont('helvetica', 'normal');
            doc.text(invoiceNumber, valueStartX, detailY);
            detailY += rowSpacing;

            // Customer Name
            doc.setFont('helvetica', 'bold');
            doc.text('Customer:', leftColX, detailY);
            doc.setFont('helvetica', 'normal');
            doc.text(invoice.customerName, valueStartX, detailY);
            detailY += rowSpacing;

            // Mobile
            doc.setFont('helvetica', 'bold');
            doc.text('Mobile:', leftColX, detailY);
            doc.setFont('helvetica', 'normal');
            doc.text(invoice.mobile, valueStartX, detailY);
            detailY += rowSpacing;

            // VAT Number
            doc.setFont('helvetica', 'bold');
            doc.text('Customer VAT:', leftColX, detailY);
            doc.setFont('helvetica', 'normal');
            doc.text(invoice.vatNumber, valueStartX, detailY);
            detailY += rowSpacing;

            // Bank Name
            doc.setFont('helvetica', 'bold');
            doc.text('Bank Name:', leftColX, detailY);
            doc.setFont('helvetica', 'normal');
            doc.text('Alrajhi Bank', valueStartX, detailY);
            detailY += rowSpacing;

            // Account Title
            doc.setFont('helvetica', 'bold');
            doc.text('Account Title:', leftColX, detailY);
            doc.setFont('helvetica', 'normal');
            doc.text('Rakan Hussein Al-Fatih Contracting Company', valueStartX, detailY);
            detailY += rowSpacing;

            // IBAN
            doc.setFont('helvetica', 'bold');
            doc.text('IBAN:', leftColX, detailY);
            doc.setFont('helvetica', 'normal');
            doc.text('SA6280000146608016555919', valueStartX, detailY);
            detailY += rowSpacing + 5;

            // Table position (start after details)
            let startY = detailY + 10;
            let startX = 15;
            const tableWidth = doc.internal.pageSize.width - 30;

            // Table column widths (sum to tableWidth)
            const colWidths = [
                tableWidth * 0.22, // Name
                tableWidth * 0.18, // Designation
                tableWidth * 0.18, // Total Hours
                tableWidth * 0.18, // Price/Hour
                tableWidth * 0.24  // Subtotal
            ];

            // Table headers
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(10);
            doc.text('Items Description', startX + 2, startY + 6, { maxWidth: colWidths[0] - 4 });
            doc.text('Division', startX + colWidths[0] + 2, startY + 6, { maxWidth: colWidths[1] - 4 });
            doc.text('Square Meters', startX + colWidths[0] + colWidths[1] + 2, startY + 6, { maxWidth: colWidths[2] - 4 });
            doc.text('Price/Sq Meter', startX + colWidths[0] + colWidths[1] + colWidths[2] + 2, startY + 6, { maxWidth: colWidths[3] - 4 });
            doc.text('Amount', startX + colWidths[0] + colWidths[1] + colWidths[2] + colWidths[3] + 2, startY + 6, { maxWidth: colWidths[4] - 4 });

            // Table content with word wrapping and truncation
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(9);
            let currentY = startY + 8;
            const pageHeight = doc.internal.pageSize.height;
            const footerHeight = 60; // Height reserved for footer
            const maxContentHeight = pageHeight - footerHeight;
            let isFirstPage = true;

            items.forEach((item, index) => {
                // Calculate row height before drawing
                const wrapText = (text, width) => {
                    let lines = doc.splitTextToSize(text, width - 4);
                    if (lines.length > 2) {
                        lines = [lines[0], lines[1].slice(0, width/2) + '...'];
                    }
                    return lines;
                };
                
                const nameLines = wrapText(item.productName || item.name || '', colWidths[0]);
                const designationLines = wrapText(item.productCode || item.code || '', colWidths[1]);
                const rowHeight = Math.max(8, nameLines.length * 7, designationLines.length * 7);

                // Check if we need a new page
                if (currentY + rowHeight > maxContentHeight) {
                    // Add a new page
                    doc.addPage();
                    isFirstPage = false;
                    
                    // Reset Y position and redraw headers
                    currentY = 30; // Start higher on continuation pages
                    doc.setFont('helvetica', 'bold');
                    doc.setFontSize(10);
                    
                    // Redraw table headers on new page
                    doc.text('Items Description', startX + 2, currentY + 6, { maxWidth: colWidths[0] - 4 });
                    doc.text('Division', startX + colWidths[0] + 2, currentY + 6, { maxWidth: colWidths[1] - 4 });
                    doc.text('Square Meters', startX + colWidths[0] + colWidths[1] + 2, currentY + 6, { maxWidth: colWidths[2] - 4 });
                    doc.text('Price/Sq Meter', startX + colWidths[0] + colWidths[1] + colWidths[2] + 2, currentY + 6, { maxWidth: colWidths[3] - 4 });
                    doc.text('Amount', startX + colWidths[0] + colWidths[1] + colWidths[2] + colWidths[3] + 2, currentY + 6, { maxWidth: colWidths[4] - 4 });
                    
                    currentY += 8; // Move below headers
                    doc.setFont('helvetica', 'normal');
                    doc.setFontSize(9);
                }

                // Draw cell borders
                let colX = startX;
                for (let i = 0; i < colWidths.length; i++) {
                    doc.rect(colX, currentY, colWidths[i], rowHeight);
                    colX += colWidths[i];
                }

                // Draw wrapped text
                nameLines.forEach((line, idx) => {
                    doc.text(line, startX + 2, currentY + 6 + (idx * 7), { maxWidth: colWidths[0] - 4 });
                });
                designationLines.forEach((line, idx) => {
                    doc.text(line, startX + colWidths[0] + 2, currentY + 6 + (idx * 7), { maxWidth: colWidths[1] - 4 });
                });
                
                // Other columns (single line, vertically centered)
                const vCenter = currentY + rowHeight / 2 + 2;
                const totalHours = `${formatNumber(item.quantity)} m`;
                const priceHour = formatNumber(item.unitPrice);
                const subtotal = formatNumber(item.subtotal);
                
                doc.text(totalHours, startX + colWidths[0] + colWidths[1] + 2, vCenter, { maxWidth: colWidths[2] - 4 });
                doc.text(priceHour, startX + colWidths[0] + colWidths[1] + colWidths[2] + 2, vCenter, { maxWidth: colWidths[3] - 4 });
                doc.text(subtotal, startX + colWidths[0] + colWidths[1] + colWidths[2] + colWidths[3] + 2, vCenter, { maxWidth: colWidths[4] - 4 });

                currentY += rowHeight;
            });

            // Add summary section on the last page
            const summaryStartY = currentY + 10;
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(10);

            // Summary box
            const summaryWidth = 80;
            const summaryX = doc.internal.pageSize.width - summaryWidth - 15;

            // Check if summary fits on current page, if not add a new page
            if (summaryStartY + 40 > maxContentHeight) {
                doc.addPage();
                currentY = 30;
            }

            // Draw summary box
            doc.rect(summaryX, currentY, summaryWidth, 30);

            // Summary text
            doc.text('Subtotal:', summaryX + 5, currentY + 8);
            doc.text('Tax (15%):', summaryX + 5, currentY + 16);
            doc.text('Total Amount:', summaryX + 5, currentY + 24);

            // Summary values
            doc.setFont('helvetica', 'normal');
            doc.text(formatCurrency(invoice.subtotal), summaryX + summaryWidth - 5, currentY + 8, { align: 'right' });
            doc.text(formatCurrency(invoice.tax), summaryX + summaryWidth - 5, currentY + 16, { align: 'right' });
            doc.text(formatCurrency(invoice.total), summaryX + summaryWidth - 5, currentY + 24, { align: 'right' });

            // Add total in words
            doc.setFont('helvetica', 'italic');
            doc.setFontSize(9);
            const totalInWords = `(${numberToWords(Math.round(invoice.total * 100) / 100)})`;
            doc.text(totalInWords, 15, currentY + 40);

            // Add page numbers to all pages
            const totalPages = doc.internal.getNumberOfPages();
            for (let i = 1; i <= totalPages; i++) {
                doc.setPage(i);
                doc.setFont('helvetica', 'normal');
                doc.setFontSize(10);
                const pageText = `Page ${i} of ${totalPages}`;
                const textWidth = doc.getStringUnitWidth(pageText) * (-80) / doc.internal.scaleFactor;
                const textX = (doc.internal.pageSize.width - textWidth) / 2;
                doc.text(pageText, textX, doc.internal.pageSize.height - 35);
            }

            // Add footer image only on the last page
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

            // After saving, clear editingFileName
            editingFileName = null;
        } catch (error) {
            console.error('Error generating invoice:', error);
            showAlert('Error generating invoice: ' + (error && error.message ? error.message : error));
            ensureInputsEditable();
        }
    } catch (error) {
        console.error('Error generating invoice:', error);
        showAlert('Error generating invoice: ' + (error && error.message ? error.message : error));
        ensureInputsEditable();
    }
});

// Helper function to convert number to words
function numberToWords(number) {
    const ones = ['', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine'];
    const tens = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];
    const teens = ['ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen'];
    const scales = ['', 'thousand', 'million', 'billion'];

    function convertGroup(n) {
        if (n === 0) return '';

        let result = '';

        // Handle hundreds
        if (n >= 100) {
            result += ones[Math.floor(n / 100)] + ' hundred ';
            n %= 100;
            if (n > 0) result += 'and ';
        }

        // Handle tens and ones
        if (n >= 20) {
            result += tens[Math.floor(n / 10)] + ' ';
            if (n % 10 > 0) result += ones[n % 10] + ' ';
        } else if (n >= 10) {
            result += teens[n - 10] + ' ';
        } else if (n > 0) {
            result += ones[n] + ' ';
        }

        return result;
    }

    if (number === 0) return 'zero';

    const wholePart = Math.floor(number);
    const decimalPart = Math.round((number - wholePart) * 100);

    let result = '';
    let groupCount = 0;
    let num = wholePart;

    while (num > 0) {
        const n = num % 1000;
        if (n !== 0) {
            const words = convertGroup(n);
            result = words + scales[groupCount] + ' ' + result;
        }
        num = Math.floor(num / 1000);
        groupCount++;
    }

    result = result.trim();

    // Add decimal part if exists
    if (decimalPart > 0) {
        result += ' and ' + convertGroup(decimalPart).trim() + 'cents';
    }

    // Capitalize first letter and ensure proper formatting
    result = result.charAt(0).toUpperCase() + result.slice(1).toLowerCase();
    
    return result || 'Zero';
}