const path = require('path');
const { app, BrowserWindow, Menu, Tray, dialog, ipcMain, nativeImage } = require('electron');

const {
    getAutomationSettings,
    saveAutomationSettings,
    exportAutomationSettings,
    importAutomationSettings,
    importServiceAccountKey,
    getServiceAccountStatus,
} = require('./sku-automation/config-store');
const { runSkuAutomation, scanConfiguredSourceFiles } = require('./sku-automation/local-automation');

const APP_ID = 'com.jbcreations.labels';

if (process.platform === 'win32') {
    app.setAppUserModelId(APP_ID);
}

const hasSingleInstanceLock = app.requestSingleInstanceLock();

if (!hasSingleInstanceLock) {
    app.quit();
}

let mainWindow = null;
let appTray = null;
let isQuitting = false;
let trayIconPromise = null;

function createTrayImageFallback() {
    const svg = `
        <svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 64">
            <rect x="8" y="8" width="48" height="48" rx="12" fill="#1f2937"/>
            <path d="M22 20h12c7 0 12 4 12 10c0 5-3 8-7 9l8 12H44l-7-10h-6v10H22V20zm9 7v8h5c3 0 5-1 5-4s-2-4-5-4h-5z" fill="#ffffff"/>
        </svg>
    `.trim();

    const image = nativeImage.createFromDataURL(`data:image/svg+xml;base64,${Buffer.from(svg).toString('base64')}`);
    return image.resize({ width: 16, height: 16 });
}

async function getTrayImage() {
    if (!trayIconPromise) {
        trayIconPromise = app.getFileIcon(process.execPath, { size: 'small' })
            .then(image => {
                if (image && !image.isEmpty()) {
                    return image.resize({ width: 16, height: 16 });
                }

                return createTrayImageFallback();
            })
            .catch(() => createTrayImageFallback());
    }

    return trayIconPromise;
}

function showMainWindow() {
    if (!mainWindow) {
        return;
    }

    if (mainWindow.isMinimized()) {
        mainWindow.restore();
    }

    mainWindow.show();
    mainWindow.focus();
}

async function ensureTray() {
    if (appTray) {
        return appTray;
    }

    appTray = new Tray(await getTrayImage());
    appTray.setToolTip('JB Creations Labels');
    appTray.setContextMenu(Menu.buildFromTemplate([
        {
            label: 'Open JB Creations Labels',
            click: () => showMainWindow(),
        },
        {
            label: 'Quit',
            click: () => {
                isQuitting = true;
                app.quit();
            },
        },
    ]));
    appTray.on('double-click', () => showMainWindow());

    return appTray;
}

function createWindow() {
    if (mainWindow) {
        showMainWindow();
        return;
    }

    mainWindow = new BrowserWindow({
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

    mainWindow.on('close', event => {
        if (isQuitting) {
            return;
        }

        event.preventDefault();
        void ensureTray();
        mainWindow.hide();
    });

    mainWindow.on('closed', () => {
        mainWindow = null;
    });

    mainWindow.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
    mainWindow.webContents.on('will-navigate', event => {
        event.preventDefault();
    });

    mainWindow.loadFile(path.join(__dirname, '..', 'index.html'));
}

app.on('second-instance', () => {
    showMainWindow();
});

app.whenReady().then(() => {
    createWindow();
    void ensureTray();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        } else {
            showMainWindow();
        }
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin' && isQuitting) {
        app.quit();
    }
});

app.on('before-quit', () => {
    isQuitting = true;

    if (appTray) {
        appTray.destroy();
        appTray = null;
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

ipcMain.handle('desktop:scan-source-files', async (_event, payload) => {
    return scanConfiguredSourceFiles(app, payload);
});