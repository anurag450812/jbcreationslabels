const fs = require('fs/promises');
const path = require('path');

const XLSX = require('xlsx');
const { google } = require('googleapis');

const {
    DEFAULT_SETTINGS,
    getAutomationSettings,
    getServiceAccountFilePath,
} = require('./config-store');

function identifyFileType(filename) {
    const filenameLower = filename.toLowerCase();

    if (filenameLower.includes('orders_') && /\d{4}-\d{2}-\d{2}/.test(filenameLower)) {
        return 'Meesho';
    }

    if (filenameLower.includes('order-csv')) {
        return 'Flipkart';
    }

    if (/^\d{10,}\.txt$/.test(filenameLower)) {
        return 'Amazon';
    }

    return 'Unknown';
}

function normalizeStringArray(value, fallback) {
    if (Array.isArray(value)) {
        return value.map(item => String(item).trim()).filter(Boolean);
    }

    if (typeof value === 'string') {
        return value.split(',').map(item => item.trim()).filter(Boolean);
    }

    return fallback;
}

function createLogger() {
    const lines = [];
    return {
        log(message) {
            lines.push(message);
        },
        all() {
            return lines;
        },
    };
}

async function fileExists(filePath) {
    try {
        await fs.access(filePath);
        return true;
    } catch {
        return false;
    }
}

async function ensureDirectory(dirPath) {
    if (!dirPath) {
        return;
    }

    await fs.mkdir(dirPath, { recursive: true });
}

function getTodayDateString() {
    const now = new Date();
    const year = now.getFullYear();
    const month = `${now.getMonth() + 1}`.padStart(2, '0');
    const day = `${now.getDate()}`.padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function isSameDate(dateA, dateB) {
    return dateA.getFullYear() === dateB.getFullYear()
        && dateA.getMonth() === dateB.getMonth()
        && dateA.getDate() === dateB.getDate();
}

function parseDelimitedText(content, delimiter) {
    const lines = content
        .split(/\r?\n/)
        .map(line => line.trim())
        .filter(Boolean);

    if (lines.length === 0) {
        return [];
    }

    const headers = lines[0].split(delimiter).map(value => value.trim());
    return lines.slice(1).map(line => {
        const values = line.split(delimiter);
        return headers.reduce((row, header, index) => {
            row[header] = (values[index] || '').trim();
            return row;
        }, {});
    });
}

async function readOrderRows(filePath, extension, logger) {
    if (extension === '.xlsx' || extension === '.xls' || extension === '.csv') {
        const workbook = XLSX.readFile(filePath, { cellDates: false, raw: false });
        const firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) {
            return [];
        }
        return XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { defval: '' });
    }

    if (extension === '.txt') {
        const text = await fs.readFile(filePath, 'utf8');
        try {
            return parseDelimitedText(text, '\t');
        } catch {
            try {
                return parseDelimitedText(text, ',');
            } catch {
                logger.log(`Warning: Could not parse structured text in ${path.basename(filePath)}.`);
                return [];
            }
        }
    }

    return [];
}

function findSkuColumn(rows, possibleSkuColumnNames) {
    if (!rows.length) {
        return null;
    }

    const columns = Object.keys(rows[0]);
    const lowerColumns = columns.map(column => column.toLowerCase());

    for (const possibleName of possibleSkuColumnNames) {
        const index = lowerColumns.indexOf(possibleName.toLowerCase());
        if (index !== -1) {
            return columns[index];
        }
    }

    return null;
}

function extractSkusFromRows(rows, fileType, settings, logger, fileName) {
    const skuColumn = findSkuColumn(rows, settings.possibleSkuColumnNames);
    if (!skuColumn) {
        logger.log(`Warning: No expected SKU column found in ${fileName}.`);
        return [];
    }

    if (fileType === 'Meesho') {
        const filterReasons = new Set(normalizeStringArray(settings.meeshoFilterReasons, DEFAULT_SETTINGS.meeshoFilterReasons));
        const creditEntryColumn = settings.meeshoCreditEntryColumn || DEFAULT_SETTINGS.meeshoCreditEntryColumn;

        return rows
            .filter(row => filterReasons.has(String(row[creditEntryColumn] || '').trim().toUpperCase()))
            .map(row => String(row[skuColumn] || '').trim())
            .filter(Boolean);
    }

    return rows
        .map(row => String(row[skuColumn] || '').trim())
        .filter(Boolean);
}

async function getFilesToProcess(payload, settings, logger) {
    const manualFiles = Array.isArray(payload.manualFilePaths)
        ? payload.manualFilePaths.filter(Boolean)
        : [];

    if (manualFiles.length > 0) {
        logger.log(`Using ${manualFiles.length} manually selected file(s).`);
        return manualFiles;
    }

    if (!settings.baseSourceFolder) {
        throw new Error('Base source folder is not configured.');
    }

    const supportedExtensions = new Set(['.xlsx', '.xls', '.csv', '.txt']);
    const items = await fs.readdir(settings.baseSourceFolder, { withFileTypes: true });
    const today = new Date();
    const files = [];

    for (const item of items) {
        if (!item.isFile()) {
            continue;
        }

        const filePath = path.join(settings.baseSourceFolder, item.name);
        const extension = path.extname(item.name).toLowerCase();
        if (!supportedExtensions.has(extension)) {
            continue;
        }

        const stats = await fs.stat(filePath);
        if (!isSameDate(stats.mtime, today)) {
            logger.log(`Skipping ${item.name} because it was not modified today.`);
            continue;
        }

        files.push(filePath);
    }

    logger.log(`Found ${files.length} source file(s) in ${settings.baseSourceFolder} modified on ${getTodayDateString()}.`);
    return files;
}

async function moveProcessedFile(sourcePath, destinationFolder, logger) {
    if (!destinationFolder) {
        return null;
    }

    await ensureDirectory(destinationFolder);
    const destinationPath = path.join(destinationFolder, path.basename(sourcePath));

    if (await fileExists(destinationPath)) {
        await fs.rm(destinationPath, { force: true });
        logger.log(`Destination file already existed and was overwritten: ${destinationPath}`);
    }

    try {
        await fs.rename(sourcePath, destinationPath);
    } catch (error) {
        if (error.code !== 'EXDEV') {
            throw error;
        }

        await fs.copyFile(sourcePath, destinationPath);
        await fs.rm(sourcePath, { force: true });
    }
    return destinationPath;
}

async function clearFolderContents(folderPath, logger) {
    if (!folderPath) {
        return [];
    }

    await ensureDirectory(folderPath);
    const items = await fs.readdir(folderPath, { withFileTypes: true });
    const deleted = [];

    for (const item of items) {
        const itemPath = path.join(folderPath, item.name);
        await fs.rm(itemPath, { recursive: true, force: true });
        deleted.push(itemPath);
    }

    logger.log(`Cleared ${deleted.length} item(s) from ${folderPath}.`);
    return deleted;
}

async function createGoogleClients(app) {
    const auth = new google.auth.GoogleAuth({
        keyFile: getServiceAccountFilePath(app),
        scopes: [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive.readonly',
        ],
    });
    const authClient = await auth.getClient();
    return {
        sheetsApi: google.sheets({ version: 'v4', auth: authClient }),
        driveApi: google.drive({ version: 'v3', auth: authClient }),
    };
}

async function getSpreadsheet(driveApi, spreadsheetName) {
    const escapedName = spreadsheetName.replace(/'/g, "\\'");
    const response = await driveApi.files.list({
        q: `mimeType='application/vnd.google-apps.spreadsheet' and name='${escapedName}' and trashed=false`,
        fields: 'files(id, name)',
        pageSize: 1,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
    });

    const file = response.data.files && response.data.files[0];
    if (!file) {
        throw new Error(`Google Sheet named "${spreadsheetName}" was not found.`);
    }

    return file;
}

async function getSheetMetadata(sheetsApi, spreadsheetId) {
    const response = await sheetsApi.spreadsheets.get({
        spreadsheetId,
        includeGridData: false,
    });
    return response.data;
}

async function updateOrdersSheet(app, settings, combinedSkus, logger) {
    const { sheetsApi, driveApi } = await createGoogleClients(app);
    const spreadsheet = await getSpreadsheet(driveApi, settings.googleSheetName);
    const spreadsheetId = spreadsheet.id;

    if (combinedSkus.length > 0) {
        const existingValues = await sheetsApi.spreadsheets.values.get({
            spreadsheetId,
            range: `${settings.ordersTabName}!A:A`,
        });
        const colValues = existingValues.data.values || [];
        const lastRowWithData = colValues.length;

        if (lastRowWithData >= 2) {
            await sheetsApi.spreadsheets.values.clear({
                spreadsheetId,
                range: `${settings.ordersTabName}!A2:A${lastRowWithData}`,
            });
            logger.log(`Cleared existing data from ${settings.ordersTabName}!A2:A${lastRowWithData}.`);
        }

        const values = combinedSkus.map(sku => [sku]);
        await sheetsApi.spreadsheets.values.update({
            spreadsheetId,
            range: `${settings.ordersTabName}!A2:A${1 + combinedSkus.length}`,
            valueInputOption: 'RAW',
            requestBody: { values },
        });
        logger.log(`Wrote ${combinedSkus.length} SKU row(s) to ${settings.ordersTabName}.`);

        const metadata = await getSheetMetadata(sheetsApi, spreadsheetId);
        const ordersSheet = (metadata.sheets || []).find(sheet => sheet.properties && sheet.properties.title === settings.ordersTabName);

        if (ordersSheet && ordersSheet.properties.gridProperties.rowCount > 1 + combinedSkus.length) {
            await sheetsApi.spreadsheets.batchUpdate({
                spreadsheetId,
                requestBody: {
                    requests: [
                        {
                            updateSheetProperties: {
                                properties: {
                                    sheetId: ordersSheet.properties.sheetId,
                                    gridProperties: {
                                        rowCount: Math.max(1 + combinedSkus.length, 2),
                                    },
                                },
                                fields: 'gridProperties.rowCount',
                            },
                        },
                    ],
                },
            });
            logger.log(`Resized ${settings.ordersTabName} to remove extra rows.`);
        }
    } else {
        logger.log('No SKUs to paste into Google Sheets. Existing ORDERS data was left unchanged.');
    }

    return { sheetsApi, spreadsheetId };
}

function toCsvString(rows) {
    return rows.map(row => row.map(value => {
        const text = String(value ?? '');
        if (/[",\n]/.test(text)) {
            return `"${text.replace(/"/g, '""')}"`;
        }
        return text;
    }).join(',')).join('\n');
}

async function exportStockAnalysis(settings, sheetsApi, spreadsheetId, logger) {
    await new Promise(resolve => setTimeout(resolve, Number(settings.waitTimeForPivotSeconds || 5) * 1000));
    const response = await sheetsApi.spreadsheets.values.get({
        spreadsheetId,
        range: `${settings.stockAnalysisTabName}!C:D`,
    });

    const values = response.data.values || [];
    const csvRows = values.length > 0 ? values : [['Column C', 'Column D']];
    const meaningfulRows = csvRows.filter((row, index) => index === 0 || (String(row[0] || '').trim() || String(row[1] || '').trim()));

    if (!settings.skuCsvPath) {
        throw new Error('SKU CSV output path is not configured.');
    }

    await ensureDirectory(path.dirname(settings.skuCsvPath));
    await fs.writeFile(settings.skuCsvPath, `${toCsvString(meaningfulRows)}\n`, 'utf8');
    logger.log(`Exported ${Math.max(meaningfulRows.length - 1, 0)} row(s) from ${settings.stockAnalysisTabName} to ${settings.skuCsvPath}.`);

    return meaningfulRows;
}

async function buildNew3Copies(settings, csvRows, logger) {
    if (!settings.sourcePdfFolder) {
        throw new Error('Source PDF folder is not configured.');
    }

    await ensureDirectory(settings.sourcePdfFolder);
    const missingSkus = [];
    const copiedFiles = [];

    for (let index = 1; index < csvRows.length; index += 1) {
        const row = csvRows[index] || [];
        const sku = String(row[0] || '').trim();
        const quantity = Number.parseInt(String(row[1] || '').trim(), 10);

        if (!sku || Number.isNaN(quantity) || quantity <= 0) {
            continue;
        }

        const sourceFile = path.join(settings.sourcePdfFolder, `${sku}.pdf`);
        if (!await fileExists(sourceFile)) {
            missingSkus.push(sku);
            logger.log(`Missing SKU PDF: ${sourceFile}`);
            continue;
        }

        for (let copyIndex = 1; copyIndex <= quantity; copyIndex += 1) {
            const destinationFile = path.join(settings.folderToClear, `${sku}_${copyIndex}.pdf`);
            await fs.copyFile(sourceFile, destinationFile);
            copiedFiles.push(destinationFile);
        }

        logger.log(`Copied ${sku} ${quantity} time(s) into ${settings.folderToClear}.`);
    }

    return {
        copiedFiles,
        missingSkus,
    };
}

async function runSkuAutomation(app, payload = {}) {
    const logger = createLogger();
    const settings = {
        ...(await getAutomationSettings(app)),
        ...(payload.settings || {}),
    };

    settings.possibleSkuColumnNames = normalizeStringArray(settings.possibleSkuColumnNames, DEFAULT_SETTINGS.possibleSkuColumnNames);
    settings.meeshoFilterReasons = normalizeStringArray(settings.meeshoFilterReasons, DEFAULT_SETTINGS.meeshoFilterReasons)
        .map(reason => reason.toUpperCase());

    const summary = {
        meesho: 0,
        flipkart: 0,
        amazon: 0,
        total: 0,
    };
    const movedFiles = [];
    const skippedFiles = [];
    const failures = [];

    if (!await fileExists(getServiceAccountFilePath(app))) {
        throw new Error('Google service account JSON has not been imported into the desktop app yet.');
    }

    if (!settings.processedFolder) {
        throw new Error('Processed/archive folder is not configured.');
    }

    if (!settings.folderToClear) {
        throw new Error('Destination new3 folder is not configured.');
    }

    const filesToProcess = await getFilesToProcess(payload, settings, logger);
    const allMeeshoSkus = [];
    const allFlipkartSkus = [];
    const allAmazonSkus = [];

    for (const filePath of filesToProcess) {
        const fileName = path.basename(filePath);
        const extension = path.extname(fileName).toLowerCase();
        const fileType = identifyFileType(fileName);

        try {
            logger.log(`Processing ${fileName} (${fileType}).`);
            const rows = await readOrderRows(filePath, extension, logger);
            if (!rows.length) {
                skippedFiles.push({ fileName, reason: 'No structured data found.' });
                logger.log(`Skipped ${fileName} because no structured rows were parsed.`);
                continue;
            }

            const currentFileSkus = extractSkusFromRows(rows, fileType, settings, logger, fileName);

            if (fileType === 'Meesho') {
                allMeeshoSkus.push(...currentFileSkus);
                summary.meesho += currentFileSkus.length;
            } else if (fileType === 'Flipkart') {
                allFlipkartSkus.push(...currentFileSkus);
                summary.flipkart += currentFileSkus.length;
            } else if (fileType === 'Amazon') {
                allAmazonSkus.push(...currentFileSkus);
                summary.amazon += currentFileSkus.length;
            } else {
                skippedFiles.push({ fileName, reason: 'Unsupported file naming pattern.' });
                logger.log(`Skipped ${fileName} because the file naming pattern was not recognized.`);
                continue;
            }

            const movedTo = await moveProcessedFile(filePath, settings.processedFolder, logger);
            if (movedTo) {
                movedFiles.push({ fileName, destinationPath: movedTo });
            }
        } catch (error) {
            failures.push({ fileName, error: error.message });
            logger.log(`Error processing ${fileName}: ${error.message}`);
        }
    }

    const combinedSkus = [...allMeeshoSkus, ...allFlipkartSkus, ...allAmazonSkus];
    summary.total = combinedSkus.length;

    const { sheetsApi, spreadsheetId } = await updateOrdersSheet(app, settings, combinedSkus, logger);
    const csvRows = await exportStockAnalysis(settings, sheetsApi, spreadsheetId, logger);
    const clearedItems = await clearFolderContents(settings.folderToClear, logger);
    const copyResult = await buildNew3Copies(settings, csvRows, logger);

    return {
        success: true,
        settings,
        summary,
        movedFiles,
        skippedFiles,
        failures,
        clearedItems,
        copiedFiles: copyResult.copiedFiles,
        missingSkus: copyResult.missingSkus,
        logs: logger.all(),
    };
}

module.exports = {
    identifyFileType,
    runSkuAutomation,
};