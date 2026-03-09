const path = require('path');
const { app, BrowserWindow, dialog, ipcMain } = require('electron');

const {
    getAutomationSettings,
    saveAutomationSettings,
    exportAutomationSettings,
    importAutomationSettings,
    importServiceAccountKey,
    getServiceAccountStatus,
} = require('./sku-automation/config-store');
const { runSkuAutomation } = require('./sku-automation/local-automation');

function createWindow() {
    const window = new BrowserWindow({
        width: 1440,
        height: 960,
        minWidth: 1180,
        minHeight: 820,
        backgroundColor: '#f8fafc',
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            contextIsolation: true,
            nodeIntegration: false,
            sandbox: false,
        },
    });

    window.loadFile(path.join(__dirname, '..', 'index.html'));
}

app.whenReady().then(() => {
    createWindow();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

ipcMain.handle('desktop:is-available', () => true);

ipcMain.handle('desktop:select-directory', async (_event, options = {}) => {
    const result = await dialog.showOpenDialog({
        title: options.title || 'Select folder',
        properties: ['openDirectory', 'createDirectory'],
    });

    if (result.canceled || result.filePaths.length === 0) {
        return null;
    }

    return result.filePaths[0];
});

ipcMain.handle('desktop:select-files', async (_event, options = {}) => {
    const result = await dialog.showOpenDialog({
        title: options.title || 'Select files',
        properties: ['openFile', 'multiSelections'],
        filters: options.filters || [
            {
                name: 'Order files',
                extensions: ['xlsx', 'xls', 'csv', 'txt'],
            },
        ],
    });

    if (result.canceled) {
        return [];
    }

    return result.filePaths;
});

ipcMain.handle('desktop:select-save-file', async (_event, options = {}) => {
    const result = await dialog.showSaveDialog({
        title: options.title || 'Select output file',
        defaultPath: options.defaultPath || 'sku.csv',
        filters: options.filters || [
            {
                name: 'CSV',
                extensions: ['csv'],
            },
        ],
    });

    if (result.canceled || !result.filePath) {
        return null;
    }

    return result.filePath;
});

ipcMain.handle('desktop:get-settings', async () => {
    const settings = await getAutomationSettings(app);
    const serviceAccount = await getServiceAccountStatus(app);

    return {
        ...settings,
        serviceAccount,
    };
});

ipcMain.handle('desktop:save-settings', async (_event, settings) => {
    return saveAutomationSettings(app, settings);
});

ipcMain.handle('desktop:export-settings', async () => {
    return exportAutomationSettings(app, dialog);
});

ipcMain.handle('desktop:import-settings', async () => {
    return importAutomationSettings(app, dialog);
});

ipcMain.handle('desktop:import-service-account', async () => {
    return importServiceAccountKey(app, dialog);
});

ipcMain.handle('desktop:run-sku-automation', async (_event, payload) => {
    return runSkuAutomation(app, payload);
});