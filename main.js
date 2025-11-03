const { app, BrowserWindow, ipcMain, shell, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
require('@electron/remote/main').initialize();

let mainWindow;

const EXPIRATION_DATE = new Date('2026-06-07T00:00:00Z'); // Temporary: set to past for testing

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1200,
        height: 800,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false,
            enableRemoteModule: true
        }
    });

    require('@electron/remote/main').enable(mainWindow.webContents);
    mainWindow.loadFile('index.html');
}

function showExpirationWindow() {
    const expWindow = new BrowserWindow({
        width: 420,
        height: 360,
        resizable: false,
        minimizable: false,
        maximizable: false,
        closable: true,
        frame: true,
        alwaysOnTop: true,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false
        }
    });
    expWindow.setMenuBarVisibility(false);
    expWindow.loadURL('data:text/html,' +
        encodeURIComponent(`
            <html>
            <head><title>Software Expired</title></head>
            <body style="font-family:sans-serif;text-align:center;padding:30px;">
                <h2 style='color:#b00;'>Software Expired</h2>
                <p>This software license has expired.<br> Please contact the help center to renew your license.</p>
                <div style='margin:20px 0;'>
                    <b>Contact:</b><br>
                    <span style='font-size:1.1em;'>Hasnain Sharif</span><br>
                    <span style='font-size:1.1em;'>+923039072713</span><br>
                    <span style='font-size:1.1em;'>hasnainsharif713@gmail.com</span>
                </div>
                <p style='color:#888;font-size:0.95em;'>Thank you for using Invoice Manager.</p>
                <button id='closeBtn' style='margin-top:20px;padding:10px 30px;font-size:1em;background:#b00;color:#fff;border:none;border-radius:5px;cursor:pointer;'>Close</button>
                <script>require('electron').remote.app.quit();document.getElementById('closeBtn').onclick = function(){ require('electron').remote.app.quit(); };</script>
            </body>
            </html>
        `)
    );
    expWindow.on('close', () => { app.quit(); });
}

// Check expiration before launching main window
if (new Date() > EXPIRATION_DATE) {
    app.whenReady().then(showExpirationWindow);
} else {
    app.whenReady().then(createWindow);
}

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
    }
});

// Storage paths
const storageDir = path.join(os.homedir(), 'AppData', 'Roaming', 'InvoiceManager');
const pdfDir = path.join(storageDir, 'PDFs');

// Create directories if they don't exist
try {
    if (!fs.existsSync(storageDir)) {
        fs.mkdirSync(storageDir, { recursive: true });
        console.log('Created storage directory:', storageDir);
    }
    if (!fs.existsSync(pdfDir)) {
        fs.mkdirSync(pdfDir, { recursive: true });
        console.log('Created PDF directory:', pdfDir);
    }
} catch (error) {
    console.error('Error creating directories:', error);
}

// IPC handlers for file operations
ipcMain.handle('save-invoice', async (event, invoice) => {
    try {
        console.log('Saving invoice data...');
        let fileName;
        if (invoice.fileName) {
            // Overwrite existing file
            fileName = invoice.fileName;
        } else {
            // Create new file
            fileName = `Invoice_${new Date().toISOString().replace(/[:.]/g, '')}.json`;
        }
        const filePath = path.join(storageDir, fileName);
        fs.writeFileSync(filePath, JSON.stringify(invoice, null, 2));
        console.log('Invoice data saved:', filePath);
        return filePath;
    } catch (error) {
        console.error('Error saving invoice:', error);
        throw new Error(`Failed to save invoice: ${error.message}`);
    }
});

ipcMain.handle('save-pdf', async (event, pdfBuffer) => {
    try {
        console.log('Saving PDF...');
        const fileName = `Invoice_${new Date().toISOString().replace(/[:.]/g, '')}.pdf`;
        const filePath = path.join(pdfDir, fileName);
        
        // Ensure the buffer is valid
        if (!Buffer.isBuffer(pdfBuffer)) {
            pdfBuffer = Buffer.from(pdfBuffer);
        }
        
        fs.writeFileSync(filePath, pdfBuffer);
        console.log('PDF saved:', filePath);
        
        // Open the PDF file
        try {
            await shell.openPath(filePath);
            console.log('PDF opened successfully');
        } catch (openError) {
            console.error('Error opening PDF:', openError);
            // Don't throw here, as the file was saved successfully
        }
        
        return filePath;
    } catch (error) {
        console.error('Error saving PDF:', error);
        throw new Error(`Failed to save PDF: ${error.message}`);
    }
});

// Add this handler after the existing IPC handlers
ipcMain.handle('open-pdf', async (event, pdfPath) => {
    try {
        await shell.openPath(pdfPath);
        console.log('PDF opened successfully:', pdfPath);
    } catch (error) {
        console.error('Error opening PDF:', error);
        throw new Error(`Failed to open PDF: ${error.message}`);
    }
}); 