const { ipcRenderer } = require('electron');
const fs = require('fs');
const path = require('path');
const os = require('os');

// Storage paths
const storageDir = path.join(os.homedir(), 'AppData', 'Roaming', 'InvoiceManager');
let currentInvoice = null;

// Load records when page loads
document.addEventListener('DOMContentLoaded', loadRecords);

// Modal elements
const modal = document.getElementById('invoiceModal');
const closeBtn = document.getElementsByClassName('close')[0];
const editInvoiceBtn = document.getElementById('editInvoiceBtn');
const deleteInvoiceBtn = document.getElementById('deleteInvoiceBtn');

// Modal event listeners
closeBtn.onclick = () => modal.style.display = 'none';
window.onclick = (event) => {
    if (event.target == modal) {
        modal.style.display = 'none';
    }
};

editInvoiceBtn.onclick = () => {
    if (currentInvoice) {
        localStorage.setItem('editInvoice', JSON.stringify(currentInvoice));
        window.location.href = 'index.html';
    }
};

deleteInvoiceBtn.onclick = () => {
    if (currentInvoice) {
        deleteInvoice(currentInvoice.fileName);
        modal.style.display = 'none';
    }
};

async function loadRecords() {
    try {
        const files = fs.readdirSync(storageDir).filter(file => file.endsWith('.json'));
        const tbody = document.getElementById('recordsTable').getElementsByTagName('tbody')[0];
        tbody.innerHTML = ''; // Clear existing records

        for (const file of files) {
            const filePath = path.join(storageDir, file);
            const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
            
            const row = tbody.insertRow();
            row.innerHTML = `
                <td>${new Date(data.date).toLocaleDateString()}</td>
                <td>${data.customerName}</td>
                <td>${formatCurrency(data.total)}</td>
                <td>
                    <button onclick="viewInvoiceDetails('${file}')" class="btn btn-primary btn-sm">View Details</button>
                    <button onclick="deleteInvoice('${file}')" class="btn btn-danger btn-sm">Delete</button>
                </td>
            `;
        }
    } catch (error) {
        console.error('Error loading records:', error);
    }
}

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

async function viewInvoiceDetails(fileName) {
    try {
        currentInvoice = {
            fileName,
            ...JSON.parse(fs.readFileSync(path.join(storageDir, fileName), 'utf8'))
        };

        const modalContent = document.querySelector('.invoice-details');
        modalContent.innerHTML = `
            <div style="display: flex; flex-direction: column; align-items: flex-start; margin-bottom: 25px;">
                <div style="display: flex; width: 100%; margin-bottom: 8px;">
                    <div style="flex: 1;"><strong>Customer:</strong> ${currentInvoice.customerName}</div>
                    <div style="flex: 1;"><strong>Mobile:</strong> ${currentInvoice.mobile}</div>
                    <div style="flex: 1;"><strong>VAT Number:</strong> ${currentInvoice.vatNumber}</div>
                </div>
                <div style="display: flex; width: 100%; margin-bottom: 8px;">
                    <div style="flex: 1;"><strong>Date:</strong> ${new Date(currentInvoice.date).toLocaleDateString()}</div>
                </div>
            </div>
            <h4 style="margin-bottom: 15px;">Items:</h4>
            <table style="margin-bottom: 25px;">
                <thead>
                    <tr>
                                                 <th style="padding: 8px 12px;">Items Description</th>
                        <th style="padding: 8px 12px;">Division</th>
                        <th style="padding: 8px 12px;">Square Meters</th>
                        <th style="padding: 8px 12px;">Price/Sq Meter</th>
                        <th style="padding: 8px 12px;">Amount</th>
                    </tr>
                </thead>
                <tbody>
                    ${currentInvoice.items.map(item => `
                        <tr>
                            <td style="padding: 8px 12px;">${item.productName}</td>
                            <td style="padding: 8px 12px;">${item.productCode}</td>
                            <td style="padding: 8px 12px;">${formatNumber(item.quantity)} m</td>
                            <td style="padding: 8px 12px;">${formatCurrency(item.unitPrice)}</td>
                            <td style="padding: 8px 12px;">${formatCurrency(item.subtotal)}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
            <div class="summary" style="margin-top: 15px;">
                <p style="margin-bottom: 8px;"><strong>Subtotal:</strong> ${formatCurrency(currentInvoice.subtotal)}</p>
                <p style="margin-bottom: 8px;"><strong>Tax (15%):</strong> ${formatCurrency(currentInvoice.tax)}</p>
                <p style="margin-bottom: 15px;"><strong>Total:</strong> ${formatCurrency(currentInvoice.total)}</p>
            </div>
        `;

        modal.style.display = 'block';
    } catch (error) {
        console.error('Error viewing invoice details:', error);
    }
}

async function deleteInvoice(fileName) {
    try {
        if (confirm('Are you sure you want to delete this invoice?')) {
            const jsonPath = path.join(storageDir, fileName);
            const pdfPath = path.join(storageDir, 'PDFs', fileName.replace('.json', '.pdf'));
            
            // Delete JSON file
            if (fs.existsSync(jsonPath)) {
                fs.unlinkSync(jsonPath);
            }
            
            // Delete PDF file if it exists
            if (fs.existsSync(pdfPath)) {
                fs.unlinkSync(pdfPath);
            }
            
            await loadRecords(); // Refresh the table
        }
    } catch (error) {
        console.error('Error deleting invoice:', error);
    }
} 