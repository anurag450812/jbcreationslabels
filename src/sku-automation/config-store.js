const fs = require('fs/promises');
const path = require('path');

const SETTINGS_FILE_NAME = 'sku-automation-settings.json';
const SERVICE_ACCOUNT_FILE_NAME = 'service-account.json';

const DEFAULT_SETTINGS = {
    baseSourceFolder: '',
    processedFolder: '',
    sourcePdfFolder: '',
    folderToClear: '',
    skuCsvPath: '',
    googleSheetName: 'order wali sheet',
    ordersTabName: 'ORDERS',
    stockAnalysisTabName: 'STOCK ANALYSIS',
    waitTimeForPivotSeconds: 5,
    possibleSkuColumnNames: ['sku', 'product id', 'item code', 'seller sku', 'order item sku'],
    meeshoCreditEntryColumn: 'Reason for Credit Entry',
    meeshoFilterReasons: ['PENDING', 'HOLD'],
};

function getConfigDir(app) {
    return path.join(app.getPath('userData'), 'sku-automation');
}

function getSettingsFilePath(app) {
    return path.join(getConfigDir(app), SETTINGS_FILE_NAME);
}

function getServiceAccountFilePath(app) {
    return path.join(getConfigDir(app), SERVICE_ACCOUNT_FILE_NAME);
}

async function ensureConfigDir(app) {
    await fs.mkdir(getConfigDir(app), { recursive: true });
}

async function readJsonFile(filePath, fallbackValue) {
    try {
        const raw = await fs.readFile(filePath, 'utf8');
        return JSON.parse(raw);
    } catch (error) {
        if (error.code === 'ENOENT') {
            return fallbackValue;
        }

        throw error;
    }
}

async function writeJsonFile(filePath, value) {
    await fs.writeFile(filePath, JSON.stringify(value, null, 2), 'utf8');
}

async function getAutomationSettings(app) {
    await ensureConfigDir(app);
    const saved = await readJsonFile(getSettingsFilePath(app), DEFAULT_SETTINGS);
    return { ...DEFAULT_SETTINGS, ...saved };
}

async function saveAutomationSettings(app, partialSettings) {
    await ensureConfigDir(app);
    const merged = {
        ...(await getAutomationSettings(app)),
        ...partialSettings,
    };

    await writeJsonFile(getSettingsFilePath(app), merged);
    return merged;
}

async function exportAutomationSettings(app, dialog) {
    const settings = await getAutomationSettings(app);
    const result = await dialog.showSaveDialog({
        title: 'Export SKU automation settings',
        defaultPath: 'jb-sku-automation-settings.json',
        filters: [
            {
                name: 'JSON',
                extensions: ['json'],
            },
        ],
    });

    if (result.canceled || !result.filePath) {
        return null;
    }

    await writeJsonFile(result.filePath, settings);
    return result.filePath;
}

async function importAutomationSettings(app, dialog) {
    const result = await dialog.showOpenDialog({
        title: 'Import SKU automation settings',
        properties: ['openFile'],
        filters: [
            {
                name: 'JSON',
                extensions: ['json'],
            },
        ],
    });

    if (result.canceled || result.filePaths.length === 0) {
        return null;
    }

    const imported = await readJsonFile(result.filePaths[0], DEFAULT_SETTINGS);
    return saveAutomationSettings(app, imported);
}

async function importServiceAccountKey(app, dialog) {
    await ensureConfigDir(app);
    const result = await dialog.showOpenDialog({
        title: 'Select Google service account JSON',
        properties: ['openFile'],
        filters: [
            {
                name: 'JSON',
                extensions: ['json'],
            },
        ],
    });

    if (result.canceled || result.filePaths.length === 0) {
        return null;
    }

    const sourcePath = result.filePaths[0];
    const destinationPath = getServiceAccountFilePath(app);
    const raw = await fs.readFile(sourcePath, 'utf8');
    JSON.parse(raw);
    await fs.writeFile(destinationPath, raw, 'utf8');

    return getServiceAccountStatus(app);
}

async function getServiceAccountStatus(app) {
    try {
        const filePath = getServiceAccountFilePath(app);
        const raw = await fs.readFile(filePath, 'utf8');
        JSON.parse(raw);
        return {
            isImported: true,
        };
    } catch (error) {
        if (error.code === 'ENOENT') {
            return {
                isImported: false,
            };
        }

        throw error;
    }
}

module.exports = {
    DEFAULT_SETTINGS,
    getAutomationSettings,
    saveAutomationSettings,
    exportAutomationSettings,
    importAutomationSettings,
    importServiceAccountKey,
    getServiceAccountStatus,
    getServiceAccountFilePath,
};