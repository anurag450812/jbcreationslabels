// Priority label list
let PRIORITY_LABELS = [
    '0baby_1',
    '0baby_1,2,11,12',
    '0baby_2,3,8,13',
    '0baby_30',
    '0baby_7,22,25,31',
    '0baby_8',
    'bd18',
    'bd21',
    'bd25',
    'bd31',
    'bd33',
    'bd34',
    'bd35',
    'bd39',
    'bd40',
    'bd46',
    'bd53',
    'bd57',
    'ch-bd58',
    'ch-bd62',
    'ch-bd64',
    'ch-bd78',
    'ch-bd84',
    'ch-fk21',
    'ch-fk28',
    'ch-fk30',
    'ch-fk45',
    'ch-fk48',
    'ch-hr31',
    'ch-ml24',
    'ch-ml26',
    'ch-ml33',
    'ch-ml38',
    'ch-ml40',
    'ch-ml41',
    'ch-ml46',
    'ch-pr01',
    'ch-pr02',
    'ch-pr03',
    'ch-rk31',
    'ch-sv21',
    'fk04',
    'gn31',
    'gn35',
    'gn42',
    'gn46',
    'hanuman01',
    'hanuman02',
    'hanuman03',
    'hr28',
    'jesus08',
    'jesus26',
    'jesus27',
    'kn13',
    'md23',
    'ml03',
    'ml05',
    'ml08',
    'ml09',
    'ml16',
    'ml17',
    'ml19',
    'ml53',
    'ml55',
    'rk028',
    'rk07',
    'rk09',
    'rk10',
    'rk15',
    'rk20',
    'rk21',
    'rk27',
    'rk29',
    'sv01',
    'sv04',
    'sv12',
    'sv19',
    'sv23',
    'sv24',
    'bd38',
    'ch-rk33',
    'hr25',
    'rk23',
    'rk69',
    'rk77',
    'rk88'
];

// Password for editing
const EDIT_PASSWORD = '200274';
const CRITERIA_PASSWORD = '200274'; // Same password for criteria settings
const DEVICE_AUTH_STORAGE_KEY = 'jb_device_authorized_v1';

// Default Label Detection Criteria
const DEFAULT_LABEL_CRITERIA = {
    meesho: {
        returnAddressPattern: /if\s+undelivered,?\s+return\s+to:?\s*([A-Za-z0-9_-]+)/i,
        returnAddressPatternString: "if\\s+undelivered,?\\s+return\\s+to:?\\s*([A-Za-z0-9_-]+)",
        returnAddressFlags: "i",
        couriers: ['DELHIVERY', 'SHADOWFAX', 'VALMO', 'XPRESS BEES'],
        pickupKeyword: 'PICKUP'
    },
    flipkart: {
        shippingAddressKeyword: 'Shipping/Customer address:',
        soldByPattern: /Sold\s*By\s*:?\s*([A-Za-z][A-Za-z0-9\s._&'()\/-]+?)\s*(?:,|\n|$)/i,
        soldByPatternString: "Sold\\s*By\\s*:?\\s*([A-Za-z][A-Za-z0-9\\s._&'()\\/-]+?)\\s*(?:,|\\n|$)",
        soldByFlags: "i"
    }
};

// Load or initialize label detection criteria
let LABEL_CRITERIA = loadLabelCriteria();

function loadLabelCriteria() {
    const stored = localStorage.getItem('labelDetectionCriteria');
    if (stored) {
        try {
            const parsed = JSON.parse(stored);
            
            // Auto-upgrade old/buggy Flipkart patterns to the current simplified version
            const CURRENT_PATTERN = DEFAULT_LABEL_CRITERIA.flipkart.soldByPatternString;
            if (parsed.flipkart && parsed.flipkart.soldByPatternString && 
                parsed.flipkart.soldByPatternString !== CURRENT_PATTERN) {
                console.log('Upgrading old Flipkart seller detection pattern to simplified version');
                parsed.flipkart.soldByPatternString = CURRENT_PATTERN;
                parsed.flipkart.soldByFlags = DEFAULT_LABEL_CRITERIA.flipkart.soldByFlags;
                localStorage.setItem('labelDetectionCriteria', JSON.stringify(parsed));
            }
            
            // Reconstruct regex patterns from strings
            if (parsed.meesho && parsed.meesho.returnAddressPatternString) {
                parsed.meesho.returnAddressPattern = new RegExp(
                    parsed.meesho.returnAddressPatternString,
                    parsed.meesho.returnAddressFlags || 'i'
                );
            }
            if (parsed.flipkart && parsed.flipkart.soldByPatternString) {
                parsed.flipkart.soldByPattern = new RegExp(
                    parsed.flipkart.soldByPatternString,
                    parsed.flipkart.soldByFlags || 'i'
                );
            }
            return parsed;
        } catch (e) {
            console.error('Error loading criteria:', e);
            return JSON.parse(JSON.stringify(DEFAULT_LABEL_CRITERIA));
        }
    }
    return JSON.parse(JSON.stringify(DEFAULT_LABEL_CRITERIA));
}

function saveLabelCriteria(criteria) {
    // Store in localStorage
    localStorage.setItem('labelDetectionCriteria', JSON.stringify(criteria));
    LABEL_CRITERIA = criteria;
    
    // Also sync to Google Sheets if configured
    syncCriteriaToGoogleSheets(criteria);
}

function resetLabelCriteriaToDefaults() {
    const defaults = JSON.parse(JSON.stringify(DEFAULT_LABEL_CRITERIA));
    saveLabelCriteria(defaults);
    return defaults;
}

async function syncCriteriaToGoogleSheets(criteria) {
    if (!GOOGLE_SHEETS_CONFIG.webAppUrl) {
        console.log('Google Sheets not configured, skipping criteria sync');
        return;
    }
    
    try {
        const response = await fetch(GOOGLE_SHEETS_CONFIG.webAppUrl, {
            method: 'POST',
            mode: 'no-cors',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                action: 'saveLabelCriteria',
                criteria: criteria
            })
        });
        console.log('Criteria synced to Google Sheets');
    } catch (error) {
        console.error('Error syncing criteria to Google Sheets:', error);
    }
}

async function loadCriteriaFromGoogleSheets() {
    if (!GOOGLE_SHEETS_CONFIG.webAppUrl) {
        return null;
    }
    
    try {
        const response = await fetch(
            `${GOOGLE_SHEETS_CONFIG.webAppUrl}?action=getLabelCriteria`,
            {
                method: 'GET',
                mode: 'cors'
            }
        );
        
        if (response.ok) {
            const data = await response.json();
            if (data.success && data.criteria) {
                return data.criteria;
            }
        }
    } catch (error) {
        console.error('Error loading criteria from Google Sheets:', error);
    }
    
    return null;
}

// Google Sheets Configuration
let GOOGLE_SHEETS_CONFIG = {
    spreadsheetId: localStorage.getItem('googleSheetsId') || '',
    sheetName: localStorage.getItem('googleSheetName') || 'STOCK COUNT',
    apiKey: localStorage.getItem('googleApiKey') || '',
    // For write access, we'll use Google Apps Script Web App
    webAppUrl: localStorage.getItem('googleWebAppUrl') || ''
};

// Global variables
let uploadedFiles = [];
let processedPDF = null;
let labelOccurrences = {}; // Store label counts globally for use during download
let pendingPasswordAction = null; // Track which action requires password ('editLabels' or 'googleSheets')
let pendingFinderDeleteEntryId = null;
let stockAlreadyDeducted = false; // Prevent duplicate stock deductions on multiple download clicks
let finderCapturedForBatch = false; // Track finder capture state for current processed batch
let currentProcessedBatchToken = '';
let lastFinderCapturedBatchToken = '';
let latestSortedPdfName = '';

// Label Counter specific variables
let counterUploadedFiles = [];

// Label Finder state
const LABEL_FINDER_DB = 'jb-label-finder-db';
const LABEL_FINDER_STORE = 'finderEntries';
const LABEL_FINDER_LIMIT = 10;
const LABEL_FINDER_CACHE = 'finder-pdf-cache-v1';
const FINDER_CAN_USE_CACHE_API = false;
let finderEntries = [];
let finderPdfCache = new Map();
let finderIndex = [];
let finderCurrentSelection = null;
let finderSearchDebounceTimer = null;
let finderPrintFrame = null;
let finderPrintUrlToRevoke = null;
let hasAppStarted = false;

// Tab switching function
function switchTab(tabName) {
    // Update tab buttons
    const tabButtons = document.querySelectorAll('.tab-btn');
    tabButtons.forEach(btn => {
        if (btn.dataset.tab === tabName) {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });

    // Update tab content
    const tabContents = document.querySelectorAll('.tab-content');
    tabContents.forEach(content => {
        if (content.id === `${tabName}-tab-content`) {
            content.classList.add('active');
        } else {
            content.classList.remove('active');
        }
    });

    if (document.body) {
        document.body.classList.toggle('finder-tab-active', tabName === 'finder');
    }
    
    // Save active tab to localStorage
    localStorage.setItem('activeTab', tabName);
}

// PDF.js worker setup
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// DOM Elements
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const processBtn = document.getElementById('processBtn');
const clearBtn = document.getElementById('clearBtn');
const statusSection = document.getElementById('statusSection');
const resultsSection = document.getElementById('resultsSection');
const progressFill = document.getElementById('progressFill');
const statusText = document.getElementById('statusText');
const downloadBtn = document.getElementById('downloadBtn');
const updateStockBtn = document.getElementById('updateStockBtn');
const priorityList = document.getElementById('priorityList');
const platformInfo = document.getElementById('platformInfo');
const platformStats = document.getElementById('platformStats');
const editLabelsBtn = document.getElementById('editLabelsBtn');
const passwordModal = document.getElementById('passwordModal');
const passwordInput = document.getElementById('passwordInput');
const submitPasswordBtn = document.getElementById('submitPasswordBtn');
const cancelPasswordBtn = document.getElementById('cancelPasswordBtn');
const passwordError = document.getElementById('passwordError');
const editLabelsModal = document.getElementById('editLabelsModal');
const labelsTextarea = document.getElementById('labelsTextarea');
const saveLabelsBtn = document.getElementById('saveLabelsBtn');
const cancelEditBtn = document.getElementById('cancelEditBtn');

// Label Counter DOM Elements
const counterUploadArea = document.getElementById('counterUploadArea');
const counterFileInput = document.getElementById('counterFileInput');
const counterProcessBtn = document.getElementById('counterProcessBtn');
const counterClearBtn = document.getElementById('counterClearBtn');
const counterFilesInfo = document.getElementById('counterFilesInfo');
const counterFilesList = document.getElementById('counterFilesList');
const counterStatusSection = document.getElementById('counterStatusSection');
const counterProgressFill = document.getElementById('counterProgressFill');
const counterStatusText = document.getElementById('counterStatusText');
const counterResultsSection = document.getElementById('counterResultsSection');
const counterTotalLabels = document.getElementById('counterTotalLabels');
const counterTotalGroups = document.getElementById('counterTotalGroups');
const counterGroupsContainer = document.getElementById('counterGroupsContainer');
const counterTextOutput = document.getElementById('counterTextOutput');
const copyCounterResultsBtn = document.getElementById('copyCounterResultsBtn');
const copyFeedback = document.getElementById('copyFeedback');

// Label Criteria Settings DOM Elements
const labelCriteriaSettingsBtn = document.getElementById('labelCriteriaSettingsBtn');
const labelCriteriaModal = document.getElementById('labelCriteriaModal');
const meeshoReturnAddressPattern = document.getElementById('meeshoReturnAddressPattern');
const meeshoCouriers = document.getElementById('meeshoCouriers');
const flipkartShippingAddressKeyword = document.getElementById('flipkartShippingAddressKeyword');
const flipkartSoldByPattern = document.getElementById('flipkartSoldByPattern');
const saveCriteriaBtn = document.getElementById('saveCriteriaBtn');
const cancelCriteriaBtn = document.getElementById('cancelCriteriaBtn');
const resetCriteriaBtn = document.getElementById('resetCriteriaBtn');

// Label Finder DOM Elements
const finderSearchInput = document.getElementById('finderSearchInput');
const finderPrintBtn = document.getElementById('finderPrintBtn');
const finderIncludeInvoiceToggle = document.getElementById('finderIncludeInvoiceToggle');
const finderSearchHelp = document.getElementById('finderSearchHelp');
const finderHistoryList = document.getElementById('finderHistoryList');
const finderResultsList = document.getElementById('finderResultsList');
const finderSelectedMeta = document.getElementById('finderSelectedMeta');
const finderLabelCanvas = document.getElementById('finderLabelCanvas');
const deviceAuthModal = document.getElementById('deviceAuthModal');
const deviceAuthInput = document.getElementById('deviceAuthInput');
const deviceAuthSubmitBtn = document.getElementById('deviceAuthSubmitBtn');
const deviceAuthError = document.getElementById('deviceAuthError');

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    initializeDeviceAccess();
});

function initializeDeviceAccess() {
    if (deviceAuthSubmitBtn) {
        deviceAuthSubmitBtn.addEventListener('click', submitDeviceAccessCode);
    }

    if (deviceAuthInput) {
        deviceAuthInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                submitDeviceAccessCode();
            }
        });
    }

    if (isDeviceAuthorized()) {
        setAppLockState(false);
        startApplicationOnce();
        return;
    }

    setAppLockState(true);
    showDeviceAccessModal();
}

function startApplicationOnce() {
    if (hasAppStarted) return;
    hasAppStarted = true;
    initializeApp();
    initCroppingTab();
}

function isDeviceAuthorized() {
    try {
        return localStorage.getItem(DEVICE_AUTH_STORAGE_KEY) === 'true';
    } catch (error) {
        console.warn('Could not read device authorization from localStorage:', error);
        return false;
    }
}

function authorizeThisDevice() {
    try {
        localStorage.setItem(DEVICE_AUTH_STORAGE_KEY, 'true');
    } catch (error) {
        console.warn('Could not persist device authorization in localStorage:', error);
    }
}

function setAppLockState(isLocked) {
    if (!document.body) return;
    document.body.classList.toggle('app-auth-locked', isLocked);
}

function showDeviceAccessModal() {
    if (!deviceAuthModal) {
        alert('Device verification is required to access this website.');
        return;
    }

    deviceAuthModal.style.display = 'flex';
    if (deviceAuthInput) {
        deviceAuthInput.value = '';
        deviceAuthInput.focus();
    }
    if (deviceAuthError) {
        deviceAuthError.style.display = 'none';
    }
}

function hideDeviceAccessModal() {
    if (deviceAuthModal) {
        deviceAuthModal.style.display = 'none';
    }
}

function submitDeviceAccessCode() {
    if (!deviceAuthInput) return;

    if (deviceAuthInput.value === EDIT_PASSWORD) {
        authorizeThisDevice();
        hideDeviceAccessModal();
        setAppLockState(false);
        startApplicationOnce();
        return;
    }

    if (deviceAuthError) {
        deviceAuthError.style.display = 'block';
    }
    deviceAuthInput.value = '';
    deviceAuthInput.focus();
}

function initializeApp() {
    // Display priority list (with default labels first)
    displayPriorityList();
    
    // Setup event listeners
    setupEventListeners();
    
    // Restore active tab from localStorage
    const savedTab = localStorage.getItem('activeTab');
    if (savedTab && (savedTab === 'sorting' || savedTab === 'counter' || savedTab === 'cropping' || savedTab === 'finder')) {
        switchTab(savedTab);
    }
    
    // Load priority labels from cloud (async, will update UI when loaded)
    loadPriorityLabelsFromCloud();
    
    // Load app configuration from cloud (async, will update settings when loaded)
    loadAppConfigFromCloud();

    // Initialize label finder state
    initializeLabelFinder();
}

function displayPriorityList() {
    priorityList.innerHTML = PRIORITY_LABELS.map(label => 
        `<div class="priority-item">${label}</div>`
    ).join('');
}

// Load priority labels from Google Sheets
async function loadPriorityLabelsFromCloud() {
    if (!GOOGLE_SHEETS_CONFIG.webAppUrl) {
        console.log('Google Sheets not configured, using default priority labels');
        return;
    }
    
    try {
        // Add action parameter to URL for GET request
        const url = `${GOOGLE_SHEETS_CONFIG.webAppUrl}?action=getPriorityLabels`;
        
        const response = await fetch(url, {
            method: 'GET',
            headers: {
                'Accept': 'application/json',
            }
        });
        
        if (response.ok) {
            const data = await response.json();
            
            if (data.success && data.labels && data.labels.length > 0) {
                PRIORITY_LABELS = data.labels;
                displayPriorityList();
                console.log(`✅ Loaded ${PRIORITY_LABELS.length} priority labels from cloud`);
                
                // Show brief notification
                showCloudSyncNotification('✅ Priority labels synced from cloud');
            } else {
                console.log('No priority labels found in cloud, using defaults');
            }
        }
    } catch (error) {
        console.error('Error loading priority labels from cloud:', error);
        // Silently use default labels if load fails
    }
}

// Load app configuration from Google Sheets
async function loadAppConfigFromCloud() {
    // First check if we have a webAppUrl in localStorage to bootstrap
    const bootstrapUrl = localStorage.getItem('googleWebAppUrl');
    if (!bootstrapUrl) {
        console.log('No Google Sheets URL configured, skipping cloud config load');
        return;
    }
    
    // Temporarily use the bootstrap URL to fetch the full config
    const tempUrl = bootstrapUrl;
    
    try {
        const url = `${tempUrl}?action=getAppConfig`;
        
        const response = await fetch(url, {
            method: 'GET',
            headers: {
                'Accept': 'application/json',
            }
        });
        
        if (response.ok) {
            const data = await response.json();
            
            if (data.success && data.config && Object.keys(data.config).length > 0) {
                // Update GOOGLE_SHEETS_CONFIG with cloud values
                if (data.config.spreadsheetId) {
                    GOOGLE_SHEETS_CONFIG.spreadsheetId = data.config.spreadsheetId;
                    localStorage.setItem('googleSheetsId', data.config.spreadsheetId);
                }
                if (data.config.sheetName) {
                    GOOGLE_SHEETS_CONFIG.sheetName = data.config.sheetName;
                    localStorage.setItem('googleSheetName', data.config.sheetName);
                }
                if (data.config.apiKey) {
                    GOOGLE_SHEETS_CONFIG.apiKey = data.config.apiKey;
                    localStorage.setItem('googleApiKey', data.config.apiKey);
                }
                if (data.config.webAppUrl) {
                    GOOGLE_SHEETS_CONFIG.webAppUrl = data.config.webAppUrl;
                    localStorage.setItem('googleWebAppUrl', data.config.webAppUrl);
                }
                
                // Update status display
                updateGoogleSheetsStatus();
                
                console.log('✅ Loaded app configuration from cloud');
                showCloudSyncNotification('✅ Google Sheets config synced from cloud');
            } else {
                console.log('No app configuration found in cloud, using local settings');
            }
        }
    } catch (error) {
        console.error('Error loading app configuration from cloud:', error);
        // Silently use local settings if load fails
    }
}

// Save app configuration to Google Sheets
async function saveAppConfigToCloud(config) {
    if (!config.webAppUrl) {
        console.log('No webAppUrl configured, skipping cloud sync');
        return;
    }
    
    try {
        const response = await fetch(config.webAppUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'text/plain',
            },
            body: JSON.stringify({
                action: 'saveAppConfig',
                config: config
            }),
            redirect: 'follow'
        });
        
        let result;
        try {
            const text = await response.text();
            result = JSON.parse(text);
        } catch (parseError) {
            if (response.ok) {
                result = { success: true };
            } else {
                throw new Error('Failed to save config: ' + response.status);
            }
        }
        
        if (result.success) {
            console.log('✅ App configuration saved to cloud');
            showCloudSyncNotification('✅ Config synced to cloud');
        }
    } catch (error) {
        console.error('Error saving app configuration to cloud:', error);
    }
}

// Show a brief notification for cloud sync status
function showCloudSyncNotification(message) {
    // Create notification element if it doesn't exist
    let notification = document.getElementById('cloudSyncNotification');
    if (!notification) {
        notification = document.createElement('div');
        notification.id = 'cloudSyncNotification';
        notification.style.cssText = `
            position: fixed;
            bottom: 20px;
            right: 20px;
            background: linear-gradient(135deg, #10b981, #059669);
            color: white;
            padding: 12px 20px;
            border-radius: 8px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.2);
            z-index: 10000;
            font-size: 14px;
            font-weight: 500;
            transform: translateY(100px);
            opacity: 0;
            transition: all 0.3s ease;
        `;
        document.body.appendChild(notification);
    }
    
    notification.textContent = message;
    
    // Show notification
    setTimeout(() => {
        notification.style.transform = 'translateY(0)';
        notification.style.opacity = '1';
    }, 100);
    
    // Hide after 3 seconds
    setTimeout(() => {
        notification.style.transform = 'translateY(100px)';
        notification.style.opacity = '0';
    }, 3000);
}

function setupEventListeners() {
    // File input
    fileInput.addEventListener('change', handleFileSelect);
    
    // Drag and drop
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });
    
    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('dragover');
    });
    
    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        const files = Array.from(e.dataTransfer.files).filter(f => f.type === 'application/pdf');
        if (files.length > 0) {
            handleFiles(files);
        }
    });
    
    // Process button
    processBtn.addEventListener('click', processLabels);
    
    // Download button
    downloadBtn.addEventListener('click', downloadSortedPDF);

    // Manual stock update button
    if (updateStockBtn) {
        updateStockBtn.addEventListener('click', updateStockToGoogleSheets);
    }
    
    // Clear button
    clearBtn.addEventListener('click', clearFiles);
    
    // Edit labels button
    editLabelsBtn.addEventListener('click', () => showPasswordModal('editLabels'));
    
    // Password modal buttons
    submitPasswordBtn.addEventListener('click', checkPassword);
    cancelPasswordBtn.addEventListener('click', hidePasswordModal);
    
    // Edit labels modal buttons
    saveLabelsBtn.addEventListener('click', saveLabels);
    cancelEditBtn.addEventListener('click', hideEditLabelsModal);
    
    // Enter key on password input
    passwordInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            checkPassword();
        }
    });
    
    // Click outside modal to close
    passwordModal.addEventListener('click', (e) => {
        if (e.target === passwordModal) {
            hidePasswordModal();
        }
    });
    
    editLabelsModal.addEventListener('click', (e) => {
        if (e.target === editLabelsModal) {
            hideEditLabelsModal();
        }
    });

    // ===============================
    // LABEL COUNTER EVENT LISTENERS
    // ===============================
    
    // Label Criteria Settings Button
    labelCriteriaSettingsBtn.addEventListener('click', () => showPasswordModal('labelCriteria'));
    
    // Criteria modal buttons
    saveCriteriaBtn.addEventListener('click', saveLabelCriteriaSettings);
    cancelCriteriaBtn.addEventListener('click', hideLabelCriteriaModal);
    resetCriteriaBtn.addEventListener('click', resetLabelCriteriaSettings);
    
    // Click outside criteria modal to close
    labelCriteriaModal.addEventListener('click', (e) => {
        if (e.target === labelCriteriaModal) {
            hideLabelCriteriaModal();
        }
    });
    
    // Counter file input
    counterFileInput.addEventListener('change', handleCounterFileSelect);
    
    // Counter drag and drop
    counterUploadArea.addEventListener('click', () => counterFileInput.click());
    
    counterUploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        counterUploadArea.classList.add('dragover');
    });
    
    counterUploadArea.addEventListener('dragleave', () => {
        counterUploadArea.classList.remove('dragover');
    });
    
    counterUploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        counterUploadArea.classList.remove('dragover');
        const files = Array.from(e.dataTransfer.files).filter(f => f.type === 'application/pdf');
        if (files.length > 0) {
            handleCounterFiles(files);
        }
    });
    
    // Counter process button
    counterProcessBtn.addEventListener('click', processCounterLabels);
    
    // Counter clear button
    counterClearBtn.addEventListener('click', clearCounterFiles);
    
    // Copy results button
    copyCounterResultsBtn.addEventListener('click', copyCounterResults);

    // Label finder
    if (finderSearchInput) {
        finderSearchInput.addEventListener('input', handleFinderSearchInput);
    }
    if (finderPrintBtn) {
        finderPrintBtn.addEventListener('click', printFinderSelection);
    }

    document.addEventListener('keydown', handleFinderGlobalKeydown);
}

// ===============================
// LABEL COUNTER FUNCTIONS
// ===============================

function handleCounterFileSelect(e) {
    const files = Array.from(e.target.files);
    if (files.length > 0) {
        handleCounterFiles(files);
    }
}

function handleCounterFiles(files) {
    // Merge with existing files instead of replacing
    const existingFileNames = new Set(counterUploadedFiles.map(f => f.name));
    const newFiles = files.filter(f => !existingFileNames.has(f.name));
    
    counterUploadedFiles = [...counterUploadedFiles, ...newFiles];
    
    if (counterUploadedFiles.length > 0) {
        counterProcessBtn.disabled = false;
        counterClearBtn.style.display = 'inline-flex';
        
        // Update upload area text
        const uploadText = counterUploadArea.querySelector('h2');
        const uploadSubtext = counterUploadArea.querySelector('p');
        uploadText.textContent = `${counterUploadedFiles.length} file(s) ready`;
        uploadSubtext.textContent = 'Drop more files or click to add more';
        
        // Show files info
        displayCounterFilesInfo();
    }
}

function displayCounterFilesInfo() {
    counterFilesInfo.style.display = 'block';
    counterFilesList.innerHTML = counterUploadedFiles.map(file => 
        `<div class="counter-file-item">📄 ${file.name}</div>`
    ).join('');
}

function clearCounterFiles() {
    counterUploadedFiles = [];
    counterFileInput.value = '';
    counterProcessBtn.disabled = true;
    counterClearBtn.style.display = 'none';
    counterFilesInfo.style.display = 'none';
    counterResultsSection.style.display = 'none';
    counterStatusSection.style.display = 'none';
    
    // Revoke skipped pages PDF URL
    if (skippedPagesPdfUrl) {
        URL.revokeObjectURL(skippedPagesPdfUrl);
        skippedPagesPdfUrl = null;
    }
    
    // Reset upload area text
    const uploadText = counterUploadArea.querySelector('h2');
    const uploadSubtext = counterUploadArea.querySelector('p');
    uploadText.textContent = 'Drop your PDF label files here';
    uploadSubtext.textContent = 'Upload labels from multiple platforms to count by courier and return address';
}

async function processCounterLabels() {
    if (counterUploadedFiles.length === 0) return;
    
    counterStatusSection.style.display = 'block';
    counterResultsSection.style.display = 'none';
    counterProgressFill.style.width = '0%';
    counterStatusText.textContent = 'Starting processing...';
    
    try {
        const allLabelsData = [];
        const skippedPages = []; // Track pages that don't match any criteria
        let totalPages = 0;
        
        for (let fileIdx = 0; fileIdx < counterUploadedFiles.length; fileIdx++) {
            const file = counterUploadedFiles[fileIdx];
            counterStatusText.textContent = `Processing file ${fileIdx + 1} of ${counterUploadedFiles.length}: ${file.name}`;
            
            const arrayBuffer = await file.arrayBuffer();
            const sourceBytes = new Uint8Array(arrayBuffer);

            // Use separate byte copies per library to avoid detached ArrayBuffer issues
            const sourcePdfJsBytes = new Uint8Array(sourceBytes);
            const sourcePdf = await pdfjsLib.getDocument({ data: sourcePdfJsBytes }).promise;
            const barcodePageNumbers = [];

            // Pass 1: identify pages that contain barcode-like content
            for (let pageNum = 1; pageNum <= sourcePdf.numPages; pageNum++) {
                const page = await sourcePdf.getPage(pageNum);
                const textContent = await page.getTextContent();
                const textVariants = extractTextWithRotationSupport(textContent);
                const rawItems = textContent.items;

                const hasBarcodeLikeContent = hasBarcodeLikeContentForCounter(textVariants, rawItems);

                if (hasBarcodeLikeContent) {
                    barcodePageNumbers.push(pageNum);
                }

                const scanProgress = ((fileIdx * 100 / counterUploadedFiles.length) +
                    (pageNum / sourcePdf.numPages * 50 / counterUploadedFiles.length));
                counterProgressFill.style.width = `${scanProgress}%`;
            }

            // No label pages in this file
            if (barcodePageNumbers.length === 0) {
                continue;
            }

            // Pass 2: create barcode-only PDF by removing non-barcode pages
            counterStatusText.textContent = `Filtering barcode pages in ${file.name}...`;
            const sourcePdfLibBytes = sourceBytes.slice();
            const sourcePdfLib = await PDFLib.PDFDocument.load(sourcePdfLibBytes);
            const filteredPdfDoc = await PDFLib.PDFDocument.create();
            const pageIndexesToKeep = barcodePageNumbers.map(pageNum => pageNum - 1);
            const copiedPages = await filteredPdfDoc.copyPages(sourcePdfLib, pageIndexesToKeep);
            for (const copiedPage of copiedPages) {
                filteredPdfDoc.addPage(copiedPage);
            }

            // Pass 3: process only pages from barcode-only PDF
            const filteredPdfBytes = await filteredPdfDoc.save();
            const barcodeOnlyPdf = await pdfjsLib.getDocument({ data: filteredPdfBytes }).promise;

            counterStatusText.textContent = `Counting barcode pages in ${file.name}...`;
            for (let filteredPageNum = 1; filteredPageNum <= barcodeOnlyPdf.numPages; filteredPageNum++) {
                const page = await barcodeOnlyPdf.getPage(filteredPageNum);
                const textContent = await page.getTextContent();
                const textVariants = extractTextWithRotationSupport(textContent);
                const rawItems = textContent.items;
                const originalPageNum = barcodePageNumbers[filteredPageNum - 1] || filteredPageNum;

                if (!hasBarcodeLikeContentForCounter(textVariants, rawItems)) {
                    continue;
                }

                // Match label info only on barcode pages
                const labelInfo = extractLabelInfoMulti(textVariants, file.name, originalPageNum, rawItems);
                if (labelInfo) {
                    allLabelsData.push(labelInfo);
                } else {
                    const combinedText = textVariants.join(' ');
                    const isInvoiceLikePage = /TAX\s+INVOICE|\bINVOICE\b|\bHSN\b|AMOUNT\s+IN\s+WORDS|TOTAL\s+AMOUNT|NET\s+AMOUNT|UNIT\s+PRICE|DESCRIPTION/i.test(combinedText);
                    if (isInvoiceLikePage) {
                        continue;
                    }
                    const reason = determineSkipReason(combinedText);
                    skippedPages.push({
                        fileName: file.name,
                        pageNum: originalPageNum,
                        fileIndex: fileIdx,
                        reason: reason
                    });
                }

                totalPages++;
                const parseProgress = ((fileIdx * 100 / counterUploadedFiles.length) +
                    (50 / counterUploadedFiles.length) +
                    (filteredPageNum / barcodeOnlyPdf.numPages * 50 / counterUploadedFiles.length));
                counterProgressFill.style.width = `${Math.min(parseProgress, 100)}%`;
            }
        }
        
        counterStatusText.textContent = 'Counting labels...';
        counterProgressFill.style.width = '100%';
        
        // Group and count labels
        const groupedResults = countLabelsByGroup(allLabelsData);
        
        // Display results
        displayCounterResults(groupedResults, totalPages, skippedPages);
        
        counterStatusSection.style.display = 'none';
        counterResultsSection.style.display = 'block';
        
    } catch (error) {
        console.error('Error processing labels:', error);
        counterStatusText.textContent = `Error: ${error.message}`;
        counterProgressFill.style.backgroundColor = '#ef4444';
    }
}

/**
 * Extract text from PDF page with rotation support.
 * Groups text items by rotation angle (0°, 90°, 180°, 270°),
 * sorts each group into proper reading order, and returns
 * an array of text variants to try for pattern matching.
 */
function extractTextWithRotationSupport(textContent) {
    const items = textContent.items;
    if (!items || items.length === 0) return [''];
    
    // Default text (simple join - works for normal non-rotated content)
    const defaultText = items.map(item => item.str).join(' ');
    
    // Collect items with position and rotation info
    const textItems = [];
    for (const item of items) {
        if (!item.str || !item.str.trim()) continue;
        
        let angle = 0;
        let x = 0, y = 0;
        
        if (item.transform && item.transform.length >= 6) {
            const [a, b, c, d, e, f] = item.transform;
            x = e;
            y = f;
            // Calculate rotation from transform matrix
            angle = Math.round(Math.atan2(b, a) * 180 / Math.PI);
            angle = ((angle % 360) + 360) % 360;
            // Snap to nearest 90°
            angle = Math.round(angle / 90) * 90 % 360;
        }
        
        textItems.push({ str: item.str, x, y, angle });
    }
    
    // Get unique rotation angles actually present in items
    const detectedAngles = [...new Set(textItems.map(item => item.angle))];
    
    // If all items are at 0° (normal), just return the default text
    if (detectedAngles.length === 1 && detectedAngles[0] === 0) return [defaultText];
    
    const texts = [defaultText];
    const LINE_THRESHOLD = 8; // pixels threshold for "same line" grouping
    
    // Build properly-ordered text for each detected rotation group
    for (const targetAngle of detectedAngles) {
        const groupItems = textItems.filter(item => item.angle === targetAngle);
        if (groupItems.length === 0) continue;
        
        const sorted = sortItemsByReadingOrder(groupItems, targetAngle, LINE_THRESHOLD);
        const text = sorted.map(item => item.str).join(' ');
        if (text.trim() && text !== defaultText) {
            texts.push(text);
        }
    }
    
    // For 90° rotated content with two spatial columns (common in Meesho labels),
    // try splitting items into left/right columns by Y-coordinate and joining each separately
    if (detectedAngles.includes(90) || detectedAngles.includes(270)) {
        const rotAngle = detectedAngles.includes(90) ? 90 : 270;
        const groupItems = textItems.filter(item => item.angle === rotAngle);
        if (groupItems.length > 4) {
            // Find the median Y to split left/right columns
            const yValues = groupItems.map(item => item.y).sort((a, b) => a - b);
            const medianY = yValues[Math.floor(yValues.length / 2)];
            const yGap = (yValues[yValues.length - 1] - yValues[0]);
            
            // Only split if there's a clear gap (columns are far apart)
            if (yGap > 200) {
                const leftCol = groupItems.filter(item => item.y < medianY);
                const rightCol = groupItems.filter(item => item.y >= medianY);
                
                if (leftCol.length > 0 && rightCol.length > 0) {
                    const leftSorted = sortItemsByReadingOrder(leftCol, rotAngle, LINE_THRESHOLD);
                    const rightSorted = sortItemsByReadingOrder(rightCol, rotAngle, LINE_THRESHOLD);
                    
                    // Left column first, then right column
                    const colText = [...leftSorted, ...rightSorted].map(item => item.str).join(' ');
                    if (colText.trim() && colText !== defaultText && !texts.includes(colText)) {
                        texts.push(colText);
                    }
                }
            }
        }
    }
    
    // Also try re-sorting ALL items assuming the entire page is at each rotation
    // This handles cases where the page rotation isn't in the transform but in layout
    for (const assumedAngle of [90, 180, 270]) {
        if (detectedAngles.length === 1 && detectedAngles[0] === assumedAngle) continue;
        
        const sorted = sortItemsByReadingOrder([...textItems], assumedAngle, LINE_THRESHOLD);
        const text = sorted.map(item => item.str).join(' ');
        if (text.trim() && text !== defaultText && !texts.includes(text)) {
            texts.push(text);
        }
    }
    
    return texts;
}

/**
 * Sort text items into reading order based on rotation angle.
 */
function sortItemsByReadingOrder(items, angle, lineThreshold) {
    return [...items].sort((a, b) => {
        switch (angle) {
            case 0: // Normal: top-to-bottom (high Y first), left-to-right
                if (Math.abs(a.y - b.y) < lineThreshold) return a.x - b.x;
                return b.y - a.y;
            case 90: // 90° CW: left-to-right (low X first), bottom-to-top
                if (Math.abs(a.x - b.x) < lineThreshold) return a.y - b.y;
                return a.x - b.x;
            case 180: // Upside down: bottom-to-top (low Y first), right-to-left
                if (Math.abs(a.y - b.y) < lineThreshold) return b.x - a.x;
                return a.y - b.y;
            case 270: // 270° CW (90° CCW): right-to-left (high X first), top-to-bottom
                if (Math.abs(a.x - b.x) < lineThreshold) return b.y - a.y;
                return b.x - a.x;
            default:
                return 0;
        }
    });
}

/**
 * Try extracting label info from multiple text variants (rotation support).
 * Falls back to item-level detection if joined text matching fails.
 */
function extractLabelInfoMulti(textVariants, fileName, pageNum, rawItems) {
    // Try each text variant (simple join, rotation-sorted, etc.)
    let bestResult = null;
    for (const text of textVariants) {
        const result = extractLabelInfo(text, fileName, pageNum);
        if (result) {
            // If we got a complete result (seller identified), return immediately
            if (result.subGroup !== 'UNKNOWN SELLER') return result;
            // Otherwise save it as fallback but keep trying
            if (!bestResult) bestResult = result;
        }
    }
    // Fallback: item-level detection (works regardless of text ordering)
    if (rawItems && rawItems.length > 0) {
        const result = extractLabelInfoFromItems(rawItems, fileName, pageNum);
        if (result) return result;
    }
    // Return the UNKNOWN SELLER result if nothing better found
    return bestResult;
}

/**
 * Try detecting platform from multiple text variants (rotation support).
 * Falls back to item-level detection if joined text matching fails.
 */
function detectPlatformFromTextMulti(textVariants, rawItems) {
    let bestResult = null;
    for (const text of textVariants) {
        const result = detectPlatformFromText(text);
        if (result.platform !== 'unknown') {
            // If seller is identified, return immediately
            if (result.subGroup !== 'UNKNOWN SELLER') return result;
            // Otherwise save as fallback
            if (!bestResult) bestResult = result;
        }
    }
    // Fallback: item-level detection
    if (rawItems && rawItems.length > 0) {
        const result = detectPlatformFromItems(rawItems);
        if (result.platform !== 'unknown') return result;
    }
    // Return best result found so far, or try combined text
    if (bestResult) return bestResult;
    const combined = textVariants.join(' ');
    return detectPlatformFromText(combined);
}

/**
 * Item-level Meesho/Flipkart detection.
 * Scans individual text items for platform patterns regardless of ordering.
 * This handles 90°/180°/270° rotated labels where joined text may be garbled.
 */
function extractLabelInfoFromItems(rawItems, fileName, pageNum) {
    // Collect all item strings (trimmed, non-empty)
    const itemStrings = rawItems
        .filter(item => item.str && item.str.trim())
        .map(item => item.str.trim());
    const allTextUpper = itemStrings.join(' ').toUpperCase();
    
    // ========================================
    // MEESHO: Look for "return to:" in individual items,
    // then find the seller ID in the NEXT item
    // ========================================
    let meeshoReturnAddress = null;
    
    for (let i = 0; i < itemStrings.length; i++) {
        const s = itemStrings[i];
        // Check if this item ends with "return to:" pattern
        const retMatch = s.match(/(?:return\s+to|undelivered.*return.*to):?\s*$/i);
        if (retMatch) {
            // Seller ID is likely the NEXT non-empty item
            // Skip single-char items (e.g., "P" from split destination codes like "DP")
            for (let j = i + 1; j < Math.min(i + 5, itemStrings.length); j++) {
                const candidate = itemStrings[j].trim();
                if (candidate && candidate.length >= 2 && /^[A-Za-z][A-Za-z0-9_-]*$/.test(candidate)) {
                    meeshoReturnAddress = candidate;
                    break;
                }
            }
            if (meeshoReturnAddress) break;
        }
        // Also check if item CONTAINS the full pattern "return to: XXXX"
        const fullMatch = s.match(/return\s+to:?\s+([A-Za-z][A-Za-z0-9_-]+)/i);
        if (fullMatch) {
            meeshoReturnAddress = fullMatch[1];
            break;
        }
        // Check for "If undelivered, return to: XXXX" across items
        const partialMatch = s.match(/if\s+undelivered.*return\s+to:?\s*([A-Za-z][A-Za-z0-9_-]+)/i);
        if (partialMatch && partialMatch[1]) {
            meeshoReturnAddress = partialMatch[1];
            break;
        }
    }
    
    if (meeshoReturnAddress) {
        // Check for courier + PICKUP across ALL items
        const couriers = LABEL_CRITERIA.meesho.couriers || ['DELHIVERY', 'SHADOWFAX', 'VALMO', 'XPRESS BEES'];
        const pickupKeyword = (LABEL_CRITERIA.meesho.pickupKeyword || 'PICKUP').toUpperCase();
        let detectedCourier = null;
        
        const hasPickup = allTextUpper.includes(pickupKeyword);
        if (hasPickup) {
            for (const courier of couriers) {
                if (allTextUpper.includes(courier.toUpperCase())) {
                    detectedCourier = courier;
                    break;
                }
            }
        }
        
        if (detectedCourier) {
            return {
                platform: 'MEESHO',
                subGroup: meeshoReturnAddress.toUpperCase(),
                courier: detectedCourier
            };
        }
    }
    
    // ========================================
    // FLIPKART: Look for "Shipping/Customer address:" and "Sold By/by" in items
    // ========================================
    const shippingKeyword = (LABEL_CRITERIA.flipkart.shippingAddressKeyword || 'Shipping/Customer address:').toUpperCase();
    const hasShippingAddr = allTextUpper.includes(shippingKeyword);
    
    if (hasShippingAddr) {
        console.log('[Flipkart Detection] Found shipping keyword, looking for seller...');
        // Look for "Sold By" or "Sold by" in individual items
        let sellerName = null;
        for (let i = 0; i < itemStrings.length; i++) {
            const s = itemStrings[i];
            // Check if item contains "Sold By: XXXX" or "Sold by : XXXX" with the name inline
            const soldByInlineMatch = s.match(/Sold\s*[Bb]y\s*:\s*([A-Za-z].+)/i);
            if (soldByInlineMatch && soldByInlineMatch[1].trim()) {
                sellerName = soldByInlineMatch[1].trim();
                console.log(`[Flipkart] Inline match at [${i}]: "${s}" -> captured: "${sellerName}"`);
                // Clean up: remove trailing address parts
                sellerName = sellerName.replace(/\s*,.*$/, '').replace(/\s+\d{2,}.*$/, '').trim();
                console.log(`[Flipkart] After cleanup: "${sellerName}"`);
                if (sellerName.length > 0) break;
            }
            // Check if item is "Sold By", "Sold By:", "Sold by :" etc. WITHOUT name inline
            // Then look at next items for the actual name
            if (/^Sold\s*[Bb]y\s*:?\s*$/i.test(s) || /^Sold\s*[Bb]y\s*$/i.test(s)) {
                console.log(`[Flipkart] Split match at [${i}]: "${s}"`);
                for (let j = i + 1; j < Math.min(i + 4, itemStrings.length); j++) {
                    let candidate = itemStrings[j].trim();
                    console.log(`[Flipkart] Checking [${j}]: "${candidate}"`);
                    // Skip colons, empty strings, short punctuation
                    if (!candidate || candidate.length < 2 || /^[:\s.,]+$/.test(candidate)) {
                        console.log(`[Flipkart] -> Skipped (empty/short/punctuation)`);
                        continue;
                    }
                    // Remove leading colon if present (e.g., ": JB CREATIONS")
                    candidate = candidate.replace(/^[:\s]+/, '').trim();
                    if (candidate.length >= 2 && /^[A-Za-z]/.test(candidate)) {
                        sellerName = candidate.replace(/\s*,.*$/, '').replace(/\s+\d{2,}.*$/, '').trim();
                        console.log(`[Flipkart] -> Accepted: "${sellerName}"`);
                        break;
                    } else {
                        console.log(`[Flipkart] -> Rejected (doesn't start with letter or too short)`);
                    }
                }
                if (sellerName) break;
            }
        }
        
        console.log(`[Flipkart] Final seller name: "${sellerName || 'UNKNOWN SELLER'}"`);
        
        return {
            platform: 'FLIPKART',
            subGroup: sellerName ? sellerName.toUpperCase() : 'UNKNOWN SELLER',
            seller: sellerName || 'Unknown'
        };
    }
    
    return null;
}

/**
 * Item-level platform detection for the sorter tab.
 */
function detectPlatformFromItems(rawItems) {
    const itemStrings = rawItems
        .filter(item => item.str && item.str.trim())
        .map(item => item.str.trim());
    const allTextUpper = itemStrings.join(' ').toUpperCase();
    
    // Meesho check: return address + courier + pickup
    let meeshoReturnAddress = null;
    for (let i = 0; i < itemStrings.length; i++) {
        const s = itemStrings[i];
        const retMatch = s.match(/(?:return\s+to|undelivered.*return.*to):?\s*$/i);
        if (retMatch) {
            for (let j = i + 1; j < Math.min(i + 5, itemStrings.length); j++) {
                const candidate = itemStrings[j].trim();
                if (candidate && candidate.length >= 2 && /^[A-Za-z][A-Za-z0-9_-]*$/.test(candidate)) {
                    meeshoReturnAddress = candidate;
                    break;
                }
            }
            if (meeshoReturnAddress) break;
        }
        const fullMatch = s.match(/return\s+to:?\s+([A-Za-z][A-Za-z0-9_-]+)/i);
        if (fullMatch) { meeshoReturnAddress = fullMatch[1]; break; }
        const partialMatch = s.match(/if\s+undelivered.*return\s+to:?\s*([A-Za-z][A-Za-z0-9_-]+)/i);
        if (partialMatch && partialMatch[1]) { meeshoReturnAddress = partialMatch[1]; break; }
    }
    
    if (meeshoReturnAddress) {
        const couriers = LABEL_CRITERIA.meesho.couriers || ['DELHIVERY', 'SHADOWFAX', 'VALMO', 'XPRESS BEES'];
        const pickupKeyword = (LABEL_CRITERIA.meesho.pickupKeyword || 'PICKUP').toUpperCase();
        const hasPickup = allTextUpper.includes(pickupKeyword);
        let hasCourier = false;
        if (hasPickup) {
            for (const courier of couriers) {
                if (allTextUpper.includes(courier.toUpperCase())) { hasCourier = true; break; }
            }
        }
        if (hasCourier) {
            return { platform: 'meesho', subGroup: meeshoReturnAddress.toUpperCase() };
        }
    }
    
    // Flipkart check
    const shippingKeyword = (LABEL_CRITERIA.flipkart.shippingAddressKeyword || 'Shipping/Customer address:').toUpperCase();
    if (allTextUpper.includes(shippingKeyword)) {
        console.log('[FK Sorter] Found shipping keyword, searching for seller...');
        let sellerName = null;
        for (let i = 0; i < itemStrings.length; i++) {
            const soldByInlineMatch = itemStrings[i].match(/Sold\s*[Bb]y\s*:\s*([A-Za-z].+)/i);
            if (soldByInlineMatch && soldByInlineMatch[1].trim()) {
                sellerName = soldByInlineMatch[1].trim().replace(/\s*,.*$/, '').replace(/\s+\d{2,}.*$/, '').trim();
                console.log(`[FK Sorter] Inline match: "${sellerName}"`);
                if (sellerName.length > 0) break;
            }
            if (/^Sold\s*[Bb]y\s*:?\s*$/i.test(itemStrings[i]) || /^Sold\s*[Bb]y\s*$/i.test(itemStrings[i])) {
                console.log(`[FK Sorter] Split match at [${i}]: "${itemStrings[i]}"`);
                for (let j = i + 1; j < Math.min(i + 4, itemStrings.length); j++) {
                    let candidate = itemStrings[j].trim();
                    if (!candidate || candidate.length < 2 || /^[:\s.,]+$/.test(candidate)) continue;
                    candidate = candidate.replace(/^[:\s]+/, '').trim();
                    if (candidate.length >= 2 && /^[A-Za-z]/.test(candidate)) {
                        sellerName = candidate.replace(/\s*,.*$/, '').replace(/\s+\d{2,}.*$/, '').trim();
                        console.log(`[FK Sorter] Split seller: "${sellerName}"`);
                        break;
                    }
                }
                if (sellerName) break;
            }
        }
        console.log(`[FK Sorter] Final: "${sellerName || 'UNKNOWN SELLER'}"`);
        return { platform: 'flipkart', subGroup: sellerName ? sellerName.toUpperCase() : 'UNKNOWN SELLER' };
    }
    
    // Amazon check
    if (allTextUpper.includes('AMAZON') || allTextUpper.includes('AMAZON.IN')) {
        return { platform: 'amazon', subGroup: 'AMAZON' };
    }
    
    return { platform: 'unknown', subGroup: 'UNKNOWN' };
}

function extractLabelInfo(text, fileName, pageNum) {
    const textUpper = text.toUpperCase();
    
    // ========================================
    // CHECK FOR MEESHO LABELS
    // Pattern: "If undelivered, return to: XXXXX" + courier with PICKUP
    // ========================================
    // Use configurable pattern
    const returnAddressPattern = LABEL_CRITERIA.meesho.returnAddressPattern;
    const returnMatch = text.match(returnAddressPattern);
    
    if (returnMatch && returnMatch[1]) {
        // This is a Meesho label - check for courier with PICKUP
        const couriers = LABEL_CRITERIA.meesho.couriers;
        const pickupKeyword = LABEL_CRITERIA.meesho.pickupKeyword;
        let detectedCourier = null;
        
        for (const courier of couriers) {
            if (textUpper.includes(courier) && textUpper.includes(pickupKeyword)) {
                detectedCourier = courier;
                break;
            }
        }
        
        if (detectedCourier) {
            return {
                platform: 'MEESHO',
                subGroup: returnMatch[1].toUpperCase(), // G-XID, XIDXID, etc.
                courier: detectedCourier
            };
        }
    }
    
    // ========================================
    // CHECK FOR FLIPKART LABELS
    // Pattern: "Shipping/Customer address:" + "Sold By: XXXXX"
    // ========================================
    const shippingKeyword = LABEL_CRITERIA.flipkart.shippingAddressKeyword;
    if (text.includes(shippingKeyword) || textUpper.includes(shippingKeyword.toUpperCase())) {
        // This is a Flipkart label - look for Sold By
        const soldByPattern = LABEL_CRITERIA.flipkart.soldByPattern;
        const soldByMatch = text.match(soldByPattern);
        
        if (soldByMatch && soldByMatch[1]) {
            const sellerName = soldByMatch[1].trim();
            return {
                platform: 'FLIPKART',
                subGroup: sellerName.toUpperCase(),
                seller: sellerName
            };
        } else {
            // Flipkart label but no seller found
            return {
                platform: 'FLIPKART',
                subGroup: 'UNKNOWN SELLER',
                seller: 'Unknown'
            };
        }
    }
    
    // No match found
    return null;
}

function determineSkipReason(text) {
    const textUpper = text.toUpperCase();
    
    // Check what patterns were found (use flexible patterns to catch rotated text)
    const hasReturnAddress = /if\s+undelivered,?\s+return\s+to:/i.test(text);
    const hasShippingAddress = text.includes('Shipping/Customer address:') || textUpper.includes('SHIPPING/CUSTOMER ADDRESS:');
    const hasCourier = textUpper.includes('DELHIVERY') || textUpper.includes('SHADOWFAX') || 
                       textUpper.includes('VALMO') || textUpper.includes('XPRESS BEES');
    const hasPickup = textUpper.includes('PICKUP');
    const hasSoldBy = /Sold\s*By\s*:/i.test(text);
    
    // Check for partial platform indicators (might be rotated/garbled)
    const hasMeeshoKeywords = textUpper.includes('MEESHO') || textUpper.includes('UNDELIVERED');
    const hasFlipkartKeywords = textUpper.includes('FLIPKART') || textUpper.includes('EKART') || 
                                 textUpper.includes('SOLD BY');
    
    // Determine most likely reason
    if (hasReturnAddress && hasCourier && !hasPickup) {
        return 'Meesho return address found, but missing "PICKUP" with courier';
    } else if (hasReturnAddress && !hasCourier) {
        return 'Meesho return address found, but no valid courier detected';
    } else if (hasShippingAddress && !hasSoldBy) {
        return 'Flipkart shipping address found, but missing "Sold By" field';
    } else if (hasShippingAddress && hasSoldBy) {
        return 'Flipkart label detected but seller name could not be extracted';
    } else if (hasCourier && !hasReturnAddress && !hasShippingAddress) {
        return 'Courier found but missing platform identifiers';
    } else if (hasMeeshoKeywords && !hasReturnAddress) {
        return 'Possible Meesho label but text may be rotated/unreadable';
    } else if (hasFlipkartKeywords && !hasShippingAddress) {
        return 'Possible Flipkart label but text may be rotated/unreadable';
    } else if (text.trim().length < 50) {
        return 'Page content too short or mostly blank';
    } else {
        return 'No Meesho or Flipkart identifiers found (may be rotated or different format)';
    }
}

function countLabelsByGroup(labelsData) {
    const results = {
        MEESHO: {
            total: 0,
            subGroups: {} // e.g., { 'XIDXID': { total: 5, couriers: { 'DELHIVERY': 3, 'SHADOWFAX': 2 } } }
        },
        FLIPKART: {
            total: 0,
            sellers: {} // e.g., { 'JB CREATIONS': 3, 'XIDLZZ': 2 }
        }
    };
    
    for (const label of labelsData) {
        if (label.platform === 'MEESHO') {
            results.MEESHO.total++;
            
            const subGroup = label.subGroup;
            if (!results.MEESHO.subGroups[subGroup]) {
                results.MEESHO.subGroups[subGroup] = {
                    total: 0,
                    couriers: {}
                };
            }
            results.MEESHO.subGroups[subGroup].total++;
            
            const courier = label.courier;
            if (!results.MEESHO.subGroups[subGroup].couriers[courier]) {
                results.MEESHO.subGroups[subGroup].couriers[courier] = 0;
            }
            results.MEESHO.subGroups[subGroup].couriers[courier]++;
            
        } else if (label.platform === 'FLIPKART') {
            results.FLIPKART.total++;
            
            const seller = label.subGroup;
            if (!results.FLIPKART.sellers[seller]) {
                results.FLIPKART.sellers[seller] = 0;
            }
            results.FLIPKART.sellers[seller]++;
        }
    }
    
    return results;
}

function displayCounterResults(groupedResults, totalPages, skippedPages) {
    const meeshoTotal = groupedResults.MEESHO.total;
    const flipkartTotal = groupedResults.FLIPKART.total;
    const grandTotal = meeshoTotal + flipkartTotal;
    
    counterTotalLabels.textContent = grandTotal;
    
    // Count unique sub-groups
    const meeshoSubGroups = Object.keys(groupedResults.MEESHO.subGroups).length;
    const flipkartSellers = Object.keys(groupedResults.FLIPKART.sellers).length;
    counterTotalGroups.textContent = meeshoSubGroups + flipkartSellers;
    
    // Generate HTML and text output
    let groupsHTML = '';
    let textOutput = '';
    
    // ========================================
    // MEESHO SECTION
    // ========================================
    if (meeshoTotal > 0) {
        groupsHTML += `
            <div class="counter-platform-section">
                <div class="platform-header meesho-header">
                    <h3>🛒 MEESHO</h3>
                    <span class="platform-total">Total: ${meeshoTotal}</span>
                </div>
        `;
        textOutput += `=== MEESHO ===\n`;
        textOutput += `MEESHO TOTAL - ${meeshoTotal}\n\n`;
        
        // Sort sub-groups: XIDXID and G-XID first, then alphabetically
        const sortedSubGroups = Object.keys(groupedResults.MEESHO.subGroups).sort((a, b) => {
            if (a === 'XIDXID') return -1;
            if (b === 'XIDXID') return 1;
            if (a === 'G-XID' || a === 'G_XID') return -1;
            if (b === 'G-XID' || b === 'G_XID') return 1;
            return a.localeCompare(b);
        });
        
        for (const subGroupName of sortedSubGroups) {
            const subGroup = groupedResults.MEESHO.subGroups[subGroupName];
            
            groupsHTML += `
                <div class="counter-group">
                    <div class="counter-group-header">
                        <h4>${subGroupName}</h4>
                        <span class="counter-group-total">TOTAL - ${subGroup.total}</span>
                    </div>
                    <div class="counter-group-items">
            `;
            
            textOutput += `${subGroupName} TOTAL - ${subGroup.total}\n`;
            
            // Sort couriers alphabetically
            const sortedCouriers = Object.keys(subGroup.couriers).sort();
            
            for (const courier of sortedCouriers) {
                const count = subGroup.couriers[courier];
                groupsHTML += `
                    <div class="counter-item">
                        <span class="courier-name">${courier}</span>
                        <span class="courier-count">${count}</span>
                    </div>
                `;
                textOutput += `  ${courier} - ${count}\n`;
            }
            
            groupsHTML += `
                    </div>
                </div>
            `;
            textOutput += '\n';
        }
        
        groupsHTML += `</div>`;
    }
    
    // ========================================
    // FLIPKART SECTION
    // ========================================
    if (flipkartTotal > 0) {
        groupsHTML += `
            <div class="counter-platform-section">
                <div class="platform-header flipkart-header">
                    <h3>🛍️ FLIPKART</h3>
                    <span class="platform-total">Total: ${flipkartTotal}</span>
                </div>
                <div class="counter-group">
                    <div class="counter-group-header flipkart-group-header">
                        <h4>Sellers</h4>
                        <span class="counter-group-total">TOTAL - ${flipkartTotal}</span>
                    </div>
                    <div class="counter-group-items">
        `;
        
        textOutput += `=== FLIPKART ===\n`;
        textOutput += `FLIPKART TOTAL - ${flipkartTotal}\n\n`;
        
        // Sort sellers alphabetically, but put JB CREATIONS first
        const sortedSellers = Object.keys(groupedResults.FLIPKART.sellers).sort((a, b) => {
            if (a.includes('JB CREATIONS') || a.includes('JB CREATION')) return -1;
            if (b.includes('JB CREATIONS') || b.includes('JB CREATION')) return 1;
            return a.localeCompare(b);
        });
        
        for (const seller of sortedSellers) {
            const count = groupedResults.FLIPKART.sellers[seller];
            groupsHTML += `
                <div class="counter-item">
                    <span class="seller-name">${seller}</span>
                    <span class="courier-count">${count}</span>
                </div>
            `;
            textOutput += `${seller} - ${count}\n`;
        }
        
        groupsHTML += `
                    </div>
                </div>
            </div>
        `;
        textOutput += '\n';
    }
    
    // ========================================
    // GRAND TOTAL
    // ========================================
    textOutput += `==============================\n`;
    textOutput += `TODAYS TOTAL - ${grandTotal}\n`;
    textOutput += `==============================\n`;
    
    // ========================================
    // SKIPPED PAGES / ERRORS
    // ========================================
    if (skippedPages.length > 0) {
        groupsHTML += `
            <div class="counter-errors-section">
                <div class="errors-header">
                    <div>
                        <h4>⚠️ Skipped Pages (${skippedPages.length})</h4>
                        <p>These pages did not match Meesho or Flipkart criteria</p>
                    </div>
                    <button class="btn btn-view-all-skipped" id="viewAllSkippedBtn" disabled>
                        📄 Creating PDF...
                    </button>
                </div>
                <div class="errors-list">
        `;
        
        textOutput += `\n⚠️ SKIPPED PAGES (${skippedPages.length}):\n`;
        
        for (const skipped of skippedPages) {
            groupsHTML += `
                <div class="error-item">
                    <span class="error-file">📄 ${skipped.fileName}</span>
                    <span class="error-page">Page ${skipped.pageNum}</span>
                    <span class="error-reason">${skipped.reason || 'Unknown criteria'}</span>
                </div>
            `;
            textOutput += `  - ${skipped.fileName}, Page ${skipped.pageNum} (${skipped.reason || 'Unknown'})\n`;
        }
        
        groupsHTML += `
                </div>
            </div>
        `;
    }
    
    counterGroupsContainer.innerHTML = groupsHTML;
    counterTextOutput.textContent = textOutput.trim();
    
    // Create PDF after DOM is rendered
    if (skippedPages.length > 0) {
        setTimeout(() => createSkippedPagesPDF(skippedPages), 100);
    }
}

// Global variable to store skipped pages PDF URL
let skippedPagesPdfUrl = null;

// Function to create a merged PDF of all skipped pages with reasons
async function createSkippedPagesPDF(skippedPages) {
    const viewBtn = document.getElementById('viewAllSkippedBtn');
    
    if (!viewBtn) {
        console.error('View button not found');
        return;
    }
    
    try {
        viewBtn.textContent = '📄 Creating PDF...';
        viewBtn.disabled = true;
        
        // Create a new PDF document
        const mergedPdfDoc = await PDFLib.PDFDocument.create();
        
        // Embed a font for adding text
        const helveticaFont = await mergedPdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
        const helveticaBold = await mergedPdfDoc.embedFont(PDFLib.StandardFonts.HelveticaBold);
        
        for (let i = 0; i < skippedPages.length; i++) {
            const skipped = skippedPages[i];
            const file = counterUploadedFiles[skipped.fileIndex];
            
            if (!file) continue;
            
            // Update button with progress
            viewBtn.textContent = `📄 Processing ${i + 1}/${skippedPages.length}...`;
            
            // Load the source PDF
            const arrayBuffer = await file.arrayBuffer();
            const sourcePdf = await PDFLib.PDFDocument.load(arrayBuffer);
            
            // Copy the specific page (pages are 0-indexed in pdf-lib)
            const [copiedPage] = await mergedPdfDoc.copyPages(sourcePdf, [skipped.pageNum - 1]);
            const addedPage = mergedPdfDoc.addPage(copiedPage);
            
            // Get page dimensions
            const { width, height } = addedPage.getSize();
            
            // Add a red banner at the top with the reason
            const bannerHeight = 60;
            const padding = 10;
            
            // Draw red background banner
            addedPage.drawRectangle({
                x: 0,
                y: height - bannerHeight,
                width: width,
                height: bannerHeight,
                color: PDFLib.rgb(0.937, 0.267, 0.267), // #ef4444
            });
            
            // Draw white text - Title (no emoji - WinAnsi encoding limitation)
            addedPage.drawText('! SKIPPED PAGE !', {
                x: padding,
                y: height - bannerHeight + 35,
                size: 14,
                font: helveticaBold,
                color: PDFLib.rgb(1, 1, 1),
            });
            
            // Draw reason text (ensure no special characters)
            const reasonText = `Reason: ${(skipped.reason || 'No Meesho/Flipkart criteria matched').replace(/[^\x00-\x7F]/g, '')}`;
            addedPage.drawText(reasonText, {
                x: padding,
                y: height - bannerHeight + 15,
                size: 10,
                font: helveticaFont,
                color: PDFLib.rgb(1, 1, 1),
            });
            
            // Draw file info
            const fileInfo = `File: ${skipped.fileName} | Page: ${skipped.pageNum}`;
            const fileInfoWidth = helveticaFont.widthOfTextAtSize(fileInfo, 8);
            addedPage.drawText(fileInfo, {
                x: width - fileInfoWidth - padding,
                y: height - bannerHeight + 25,
                size: 8,
                font: helveticaFont,
                color: PDFLib.rgb(1, 1, 1),
            });
            
            // Allow UI to update
            await new Promise(resolve => setTimeout(resolve, 0));
        }
        
        // Save the merged PDF
        viewBtn.textContent = '💾 Saving PDF...';
        const mergedPdfBytes = await mergedPdfDoc.save();
        const blob = new Blob([mergedPdfBytes], { type: 'application/pdf' });
        
        // Revoke old URL if exists
        if (skippedPagesPdfUrl) {
            URL.revokeObjectURL(skippedPagesPdfUrl);
        }
        
        // Create new URL
        skippedPagesPdfUrl = URL.createObjectURL(blob);
        
        // Update button to view the PDF
        viewBtn.textContent = `👁️ View All ${skippedPages.length} Skipped Pages`;
        viewBtn.disabled = false;
        viewBtn.onclick = () => {
            window.open(skippedPagesPdfUrl, '_blank');
        };
        
    } catch (error) {
        console.error('Error creating skipped pages PDF:', error);
        if (viewBtn) {
            viewBtn.textContent = '❌ Error creating PDF';
            viewBtn.disabled = true;
        }
        alert('Error creating skipped pages PDF: ' + error.message);
    }
}

function copyCounterResults() {
    const text = counterTextOutput.textContent;
    navigator.clipboard.writeText(text).then(() => {
        copyFeedback.style.display = 'inline';
        setTimeout(() => {
            copyFeedback.style.display = 'none';
        }, 2000);
    }).catch(err => {
        console.error('Failed to copy:', err);
        alert('Failed to copy results. Please select and copy manually.');
    });
}

// ===============================
// ORIGINAL LABEL SORTING FUNCTIONS
// ===============================

function clearFiles() {
    uploadedFiles = [];
    fileInput.value = '';
    processBtn.disabled = true;
    clearBtn.style.display = 'none';
    platformInfo.style.display = 'none';
    resultsSection.style.display = 'none';
    statusSection.style.display = 'none';
    
    // Reset stock deduction flag for new batch
    stockAlreadyDeducted = false;
    finderCapturedForBatch = false;
    currentProcessedBatchToken = '';
    labelOccurrences = {};
    
    // Reset upload area text
    const uploadText = uploadArea.querySelector('h2');
    const uploadSubtext = uploadArea.querySelector('p');
    uploadText.textContent = 'Drop your PDF files here';
    uploadSubtext.textContent = 'or click to browse (Ctrl+Click or Shift+Click to select multiple)';
}

function handleFileSelect(e) {
    const files = Array.from(e.target.files);
    if (files.length > 0) {
        handleFiles(files);
    }
}

async function handleFiles(files) {
    // Merge with existing files instead of replacing
    const existingFileNames = new Set(uploadedFiles.map(f => f.name));
    const newFiles = files.filter(f => !existingFileNames.has(f.name));
    
    // Add only new files to avoid duplicates
    uploadedFiles = [...uploadedFiles, ...newFiles];
    
    if (uploadedFiles.length > 0) {
        processBtn.disabled = false;
        clearBtn.style.display = 'block';
        
        // Update upload area text
        const uploadText = uploadArea.querySelector('h2');
        const uploadSubtext = uploadArea.querySelector('p');
        uploadText.textContent = `${uploadedFiles.length} file(s) selected`;
        
        // Show file names (limit display to avoid overflow)
        const displayNames = uploadedFiles.slice(0, 3).map(f => f.name);
        if (uploadedFiles.length > 3) {
            displayNames.push(`... and ${uploadedFiles.length - 3} more`);
        }
        uploadSubtext.textContent = displayNames.join(', ');
        
        // Get file info with page counts
        const fileInfoList = await detectPlatforms(uploadedFiles);
        
        // Display file info
        displayPlatformInfo(fileInfoList);
    }
}

async function detectPlatforms(files) {
    // Get page counts for each file
    const fileInfoPromises = files.map(async (file) => {
        try {
            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            return {
                name: file.name,
                pageCount: pdf.numPages
            };
        } catch (error) {
            console.error(`Error reading ${file.name}:`, error);
            return {
                name: file.name,
                pageCount: 0
            };
        }
    });
    
    return await Promise.all(fileInfoPromises);
}

function displayPlatformInfo(fileInfoList) {
    let html = '';
    
    fileInfoList.forEach(fileInfo => {
        html += `
            <div class="platform-stat">
                <div class="platform-stat-icon">📄</div>
                <div class="platform-stat-content">
                    <span class="platform-stat-name">${fileInfo.name}</span>
                    <span class="platform-stat-count">${fileInfo.pageCount} page${fileInfo.pageCount !== 1 ? 's' : ''}</span>
                </div>
            </div>
        `;
    });
    
    platformStats.innerHTML = html;
    platformInfo.style.display = 'block';
}

async function processLabels() {
    if (uploadedFiles.length === 0) return;
    
    // Show status section
    statusSection.style.display = 'block';
    resultsSection.style.display = 'none';
    processBtn.disabled = true;
    
    try {
        updateProgress(0, 'Starting processing...');
        
        // Process all PDFs
        const allPages = [];
        let fileIndex = 0;
        
        for (const file of uploadedFiles) {
            fileIndex++;
            updateProgress(
                (fileIndex / uploadedFiles.length) * 50,
                `Processing ${file.name} (${fileIndex}/${uploadedFiles.length})...`
            );
            
            const pages = await extractPagesFromPDF(file);
            allPages.push(...pages);
        }
        
        updateProgress(60, 'Analyzing labels...');
        
        // Sort pages based on priority
        const { priorityPages, otherPages, outputPages, matchedLabels, labelCounts, labelStatPages } = sortPages(allPages);
        
        // Store label counts globally for use during download
        labelOccurrences = labelCounts;
        
        // Reset stock deduction flag for this new processing batch
        stockAlreadyDeducted = false;
        finderCapturedForBatch = false;
        currentProcessedBatchToken = `batch_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
        
        updateProgress(80, 'Creating sorted PDF...');
        
        // Create new PDF with sorted pages
        processedPDF = await createSortedPDF(outputPages, uploadedFiles);
        latestSortedPdfName = `sorted_labels_${new Date().getTime()}.pdf`;
        
        updateProgress(100, 'Complete!');
        
        // Display results with barcode-label-only counts
        const totalLabelPages = labelStatPages.length;
        const matchedLabelPages = priorityPages.length;
        const unmatchedLabelPages = Math.max(0, totalLabelPages - matchedLabelPages);
        displayResults(totalLabelPages, matchedLabelPages, unmatchedLabelPages, matchedLabels, labelStatPages, labelCounts);
        
    } catch (error) {
        console.error('Error processing PDFs:', error);
        statusText.textContent = `Error: ${error.message}`;
        statusText.style.color = 'red';
        processBtn.disabled = false;
    }
}

async function extractPagesFromPDF(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const pages = [];
    
    // Detect platform from filename as fallback
    const filenamePlatform = detectPlatformFromFilename(file.name);
    
    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        
        // Extract text with rotation support (handles 90°, 180°, 270° rotated labels)
        const textVariants = extractTextWithRotationSupport(textContent);
        
        // Detect platform from text content trying all rotation variants
        // Pass raw items for item-level fallback detection (handles rotated labels)
        const rawItems = textContent.items;
        const textDetection = detectPlatformFromTextMulti(textVariants, rawItems);
        
        // Use text-based detection if available, otherwise use filename
        const platform = textDetection.platform !== 'unknown' ? textDetection.platform : filenamePlatform;
        const subGroup = textDetection.subGroup;
        
        // Combine all text variants for priority label search (handles rotated text)
        const combinedText = textVariants.join(' ');
        
        pages.push({
            fileName: file.name,
            pageNumber: i,
            text: combinedText.toLowerCase(),
            platform: platform,
            subGroup: subGroup,  // Store sub-group info (XID, JBCREATIONS, seller name, etc.)
            fileRef: file  // Store file reference instead of ArrayBuffer
        });
    }
    
    return pages;
}

function detectPlatformFromFilename(filename) {
    const name = filename.toLowerCase();
    if (name.includes('flipkart')) return 'flipkart';
    if (name.includes('amazon')) return 'amazon';
    if (name.includes('meesho')) return 'meesho';
    return 'unknown';
}

/**
 * Detect platform from page text content (more accurate than filename)
 * Uses same logic as label counter section with configurable criteria
 */
function detectPlatformFromText(text) {
    const textUpper = text.toUpperCase();
    
    // CHECK FOR MEESHO LABELS
    // Pattern: "If undelivered, return to: XXXXX" + courier with PICKUP
    const returnAddressPattern = LABEL_CRITERIA.meesho.returnAddressPattern;
    const returnMatch = text.match(returnAddressPattern);
    
    if (returnMatch && returnMatch[1]) {
        // Check for courier with PICKUP to confirm it's Meesho
        const couriers = LABEL_CRITERIA.meesho.couriers;
        const pickupKeyword = LABEL_CRITERIA.meesho.pickupKeyword;
        let hasCourier = false;
        
        for (const courier of couriers) {
            if (textUpper.includes(courier) && textUpper.includes(pickupKeyword)) {
                hasCourier = true;
                break;
            }
        }
        
        if (hasCourier) {
            const returnAddress = returnMatch[1].toUpperCase();
            
            // Identify sub-group: G-XID, XIDXID, JBCREATIONS, etc.
            if (returnAddress.includes('XID') || returnAddress === 'XIDXID') {
                return { platform: 'meesho', subGroup: returnAddress };
            } else if (returnAddress.includes('JB') || returnAddress.includes('JBCREATION')) {
                return { platform: 'meesho', subGroup: returnAddress };
            } else if (returnAddress.startsWith('G-')) {
                return { platform: 'meesho', subGroup: returnAddress };
            } else {
                return { platform: 'meesho', subGroup: returnAddress };
            }
        }
    }
    
    // CHECK FOR FLIPKART LABELS
    // Pattern: "Shipping/Customer address:" + "Sold By: XXXXX"
    const shippingKeyword = LABEL_CRITERIA.flipkart.shippingAddressKeyword;
    if (text.includes(shippingKeyword) || textUpper.includes(shippingKeyword.toUpperCase())) {
        const soldByPattern = LABEL_CRITERIA.flipkart.soldByPattern;
        const soldByMatch = text.match(soldByPattern);
        
        if (soldByMatch && soldByMatch[1]) {
            const sellerName = soldByMatch[1].trim().toUpperCase();
            
            // Check if it's JB CREATIONS or similar
            if (sellerName.includes('JB CREATION') || sellerName.includes('JB CREATION')) {
                return { platform: 'flipkart', subGroup: sellerName };
            } else {
                return { platform: 'flipkart', subGroup: sellerName };
            }
        } else {
            return { platform: 'flipkart', subGroup: 'UNKNOWN SELLER' };
        }
    }
    
    // CHECK FOR AMAZON (basic detection)
    if (textUpper.includes('AMAZON') || textUpper.includes('AMAZON.IN')) {
        return { platform: 'amazon', subGroup: 'AMAZON' };
    }
    
    // No platform detected
    return { platform: 'unknown', subGroup: 'UNKNOWN' };
}

function sortPages(pages) {
    const priorityPages = [];
    const otherPages = [];
    const matchedLabels = new Set();
    const priorityMap = new Map();
    const labelCounts = {}; // Count occurrences of each label
    
    // Initialize label counts
    PRIORITY_LABELS.forEach(label => {
        labelCounts[label] = 0;
    });
    
    // Create priority index map
    PRIORITY_LABELS.forEach((label, index) => {
        priorityMap.set(label.toLowerCase(), index);
    });
    
    // Step 1: pair label + invoice pages first (same file, immediate next page)
    const pageOrders = buildPageOrdersByAdjacency(pages);

    // Step 2: sort based on priority using only the LABEL page text in each order
    const priorityOrders = [];
    const otherOrders = [];

    for (const order of pageOrders) {
        const page = order.labelPage;

        // Only barcode-identified label pages participate in label matching/counting stats
        if (!order.isBarcodeLabel) {
            otherOrders.push(order);
            continue;
        }

        let matched = false;
        let matchedLabel = null;
        let priorityIndex = Infinity;

        for (const label of PRIORITY_LABELS) {
            const labelLower = label.toLowerCase();
            const textLower = page.text || '';
            const index = textLower.indexOf(labelLower);

            if (index !== -1) {
                const charAfterLabel = textLower.charAt(index + labelLower.length);
                if (charAfterLabel !== ',' && !/\d/.test(charAfterLabel)) {
                    matched = true;
                    const currentIndex = priorityMap.get(labelLower);
                    if (currentIndex < priorityIndex) {
                        priorityIndex = currentIndex;
                        matchedLabel = label;
                    }
                }
            }
        }

        if (matched) {
            const enrichedLabelPage = { ...page, priorityIndex, matchedLabel };
            priorityPages.push(enrichedLabelPage);
            priorityOrders.push({
                ...order,
                labelPage: enrichedLabelPage,
                priorityIndex,
                matchedLabel,
            });
            matchedLabels.add(matchedLabel);
            labelCounts[matchedLabel] = (labelCounts[matchedLabel] || 0) + 1;
        } else {
            otherPages.push(page);
            otherOrders.push(order);
        }
    }

    priorityOrders.sort((a, b) => a.priorityIndex - b.priorityIndex);

    const outputPages = [];
    for (const order of [...priorityOrders, ...otherOrders]) {
        outputPages.push(order.labelPage);
        for (const invoicePage of order.invoicePages) {
            outputPages.push(invoicePage);
        }
    }

    const labelStatPages = pageOrders
        .filter(order => order.isBarcodeLabel)
        .map(order => order.labelPage);
    
    // Filter to only labels that were found
    const foundLabelCounts = {};
    for (const label of matchedLabels) {
        foundLabelCounts[label] = labelCounts[label];
    }
    
    return {
        priorityPages,
        otherPages,
        outputPages,
        labelStatPages,
        matchedLabels: Array.from(matchedLabels),
        labelCounts: foundLabelCounts // Return the counts
    };
}

function extractOverlayNumberFromText(text) {
    if (!text || typeof text !== 'string') return null;
    const match = text.match(/#\s*(\d{1,8})\b/);
    return match ? match[1] : null;
}

function hasBarcodeLikeTextForSorter(text) {
    if (!text || typeof text !== 'string') return false;
    const upper = text.toUpperCase();

    if (upper.includes('BARCODE')) return true;
    if (/\*[A-Z0-9\-]{6,}\*/.test(upper)) return true;

    // Compact alphanumeric barcode-like tokens
    const tokens = upper.split(/\s+/).filter(Boolean);
    for (const token of tokens) {
        const compact = token.replace(/[^A-Z0-9\-]/g, '');
        if (compact.length >= 10 && /\d/.test(compact) && /[A-Z]/.test(compact)) {
            return true;
        }
        if (/^\d{12,}$/.test(compact)) {
            return true;
        }
    }

    return false;
}

function hasBarcodeLikeContentForCounter(textVariants, rawItems) {
    const variants = Array.isArray(textVariants) ? textVariants : [];
    const combinedText = variants.join(' ');
    const combinedUpper = combinedText.toUpperCase();

    const hasInvoiceHeavyMarkers =
        /TAX\s+INVOICE|\bINVOICE\b|\bHSN\b|AMOUNT\s+IN\s+WORDS|TOTAL\s+AMOUNT|NET\s+AMOUNT/i.test(combinedText);

    const hasStrongBarcodeKeywords =
        /\bBARCODE\b|\bAWB\b|TRACKING\s*(ID|NO|NUMBER)?/i.test(combinedText);

    const hasStarWrappedBarcode = /\*[A-Z0-9\-]{6,}\*/.test(combinedUpper);

    const tokens = (rawItems || [])
        .map(item => (item && item.str ? String(item.str).trim().toUpperCase() : ''))
        .filter(Boolean);

    let hasLongMixedToken = false;
    let longNumericTokenCount = 0;

    for (const token of tokens) {
        const compact = token.replace(/[^A-Z0-9\-]/g, '');
        if (!compact) continue;

        if (compact.length >= 10 && /[A-Z]/.test(compact) && /\d/.test(compact)) {
            hasLongMixedToken = true;
        }
        if (/^\d{12,}$/.test(compact)) {
            longNumericTokenCount++;
        }
    }

    const hasShippingLabelMarkers =
        /SHIPPING|DELIVERY|PICKUP|SOLD\s*BY|RETURN\s+TO|SHIP\s+TO|CUSTOMER\s+ADDRESS|AWB\s*NO/i.test(combinedText);

    if (hasStrongBarcodeKeywords || hasStarWrappedBarcode) {
        return true;
    }

    // Mixed long tokens can appear in invoices (SKU/order ids), so require label context when invoice markers exist
    if (hasLongMixedToken) {
        if (hasInvoiceHeavyMarkers && !hasShippingLabelMarkers) return false;
        return true;
    }

    if (longNumericTokenCount > 0) {
        if (hasInvoiceHeavyMarkers) return false;
        return hasShippingLabelMarkers;
    }

    return false;
}

function buildPageOrdersByAdjacency(pages) {
    const orders = [];
    const usedKeys = new Set();

    // Group by source file first
    const byFile = new Map();
    for (const page of pages) {
        if (!byFile.has(page.fileName)) byFile.set(page.fileName, []);
        byFile.get(page.fileName).push(page);
    }

    for (const filePages of byFile.values()) {
        filePages.sort((a, b) => a.pageNumber - b.pageNumber);

        for (let i = 0; i < filePages.length; i++) {
            const currentPage = filePages[i];
            const currentKey = `${currentPage.fileName}::${currentPage.pageNumber}`;
            if (usedKeys.has(currentKey)) continue;

            const isLabelPage = hasBarcodeLikeTextForSorter(currentPage.text || '');
            if (!isLabelPage) continue;

            const nextPage = filePages[i + 1];
            if (nextPage) {
                const nextKey = `${nextPage.fileName}::${nextPage.pageNumber}`;
                if (!usedKeys.has(nextKey)) {
                    orders.push({ labelPage: currentPage, invoicePages: [nextPage], isBarcodeLabel: true });
                    usedKeys.add(currentKey);
                    usedKeys.add(nextKey);
                    i++; // skip the paired next page
                    continue;
                }
            }

            // Label page exists but no available immediate next page to pair
            orders.push({ labelPage: currentPage, invoicePages: [], isBarcodeLabel: true });
            usedKeys.add(currentKey);
        }
    }

    // Add any unpaired pages as standalone orders (keeps all pages in output)
    const remainingPages = [...pages].sort((a, b) => {
        if (a.fileName === b.fileName) return a.pageNumber - b.pageNumber;
        return a.fileName.localeCompare(b.fileName);
    });

    for (const page of remainingPages) {
        const key = `${page.fileName}::${page.pageNumber}`;
        if (!usedKeys.has(key)) {
            orders.push({ labelPage: page, invoicePages: [], isBarcodeLabel: false });
            usedKeys.add(key);
        }
    }

    return orders;
}

async function createSortedPDF(sortedPages, originalFiles) {
    const { PDFDocument } = PDFLib;
    const mergedPdf = await PDFDocument.create();
    
    // Create a map of filename to file object
    const fileMap = new Map();
    originalFiles.forEach(file => {
        fileMap.set(file.name, file);
    });
    
    // Group pages by source file
    const fileGroups = new Map();
    for (const page of sortedPages) {
        if (!fileGroups.has(page.fileName)) {
            fileGroups.set(page.fileName, []);
        }
        fileGroups.get(page.fileName).push(page);
    }
    
    // Load each unique source PDF once
    const loadedPdfs = new Map();
    
    for (const [fileName, pages] of fileGroups.entries()) {
        try {
            const file = fileMap.get(fileName);
            if (file) {
                const arrayBuffer = await file.arrayBuffer();
                const sourcePdf = await PDFDocument.load(arrayBuffer);
                loadedPdfs.set(fileName, sourcePdf);
            }
        } catch (error) {
            console.error(`Error loading PDF ${fileName}:`, error);
        }
    }
    
    // Copy pages in sorted order
    for (const page of sortedPages) {
        try {
            const sourcePdf = loadedPdfs.get(page.fileName);
            if (sourcePdf) {
                const [copiedPage] = await mergedPdf.copyPages(sourcePdf, [page.pageNumber - 1]);
                mergedPdf.addPage(copiedPage);
            }
        } catch (error) {
            console.error(`Error copying page ${page.pageNumber} from ${page.fileName}:`, error);
        }
    }
    
    return mergedPdf;
}

function displayResults(total, matched, unmatched, labels, allPages, labelCounts = {}) {
    // Update stats
    document.getElementById('totalPages').textContent = total;
    document.getElementById('matchedPages').textContent = matched;
    document.getElementById('unmatchedPages').textContent = unmatched;
    
    // Calculate comprehensive platform breakdown with sub-groups
    const platformCounts = {
        flipkart: 0,
        amazon: 0,
        meesho: 0,
        unknown: 0
    };
    
    const subGroupCounts = {
        flipkart: {},
        meesho: {},
        amazon: {}
    };
    
    allPages.forEach(page => {
        if (platformCounts.hasOwnProperty(page.platform)) {
            platformCounts[page.platform]++;
            
            // Count sub-groups (XID, JB CREATIONS, seller names, etc.)
            if (page.subGroup && page.subGroup !== 'UNKNOWN') {
                if (!subGroupCounts[page.platform][page.subGroup]) {
                    subGroupCounts[page.platform][page.subGroup] = 0;
                }
                subGroupCounts[page.platform][page.subGroup]++;
            }
        }
    });
    
    // Display comprehensive platform breakdown
    const platformBreakdownStats = document.getElementById('platformBreakdownStats');
    let breakdownHtml = '';
    
    // MEESHO Section with sub-groups
    if (platformCounts.meesho > 0) {
        breakdownHtml += `<div class="platform-section meesho-section">`;
        breakdownHtml += `
            <div class="platform-breakdown-item platform-breakdown-main">
                <div class="platform-breakdown-icon">🛍️</div>
                <span class="platform-breakdown-name">Meesho</span>
                <span class="platform-breakdown-count">${platformCounts.meesho}</span>
                <span class="platform-breakdown-label">labels processed</span>
            </div>
        `;
        
        // Show sub-groups (XID, JBCREATIONS, G-XID, etc.)
        const meeshoSubGroups = Object.entries(subGroupCounts.meesho);
        if (meeshoSubGroups.length > 0) {
            // Sort sub-groups: JB-related first, then alphabetically
            meeshoSubGroups.sort((a, b) => {
                const aIsJB = a[0].includes('JB') || a[0].includes('JBCREATION');
                const bIsJB = b[0].includes('JB') || b[0].includes('JBCREATION');
                if (aIsJB && !bIsJB) return -1;
                if (!aIsJB && bIsJB) return 1;
                return a[0].localeCompare(b[0]);
            });
            
            meeshoSubGroups.forEach(([subGroup, count]) => {
                breakdownHtml += `
                    <div class="platform-breakdown-item platform-breakdown-sub">
                        <div class="platform-breakdown-icon">↳</div>
                        <span class="platform-breakdown-name">${subGroup}</span>
                        <span class="platform-breakdown-count">${count}</span>
                        <span class="platform-breakdown-label">labels</span>
                    </div>
                `;
            });
        }
        breakdownHtml += `</div>`; // Close meesho-section
    }
    
    // FLIPKART Section with sellers
    if (platformCounts.flipkart > 0) {
        breakdownHtml += `<div class="platform-section flipkart-section">`;
        breakdownHtml += `
            <div class="platform-breakdown-item platform-breakdown-main">
                <div class="platform-breakdown-icon">🛒</div>
                <span class="platform-breakdown-name">Flipkart</span>
                <span class="platform-breakdown-count">${platformCounts.flipkart}</span>
                <span class="platform-breakdown-label">labels processed</span>
            </div>
        `;
        
        // Show sellers
        const flipkartSellers = Object.entries(subGroupCounts.flipkart);
        if (flipkartSellers.length > 0) {
            // Sort sellers: JB CREATIONS first, then alphabetically
            flipkartSellers.sort((a, b) => {
                const aIsJB = a[0].includes('JB CREATION');
                const bIsJB = b[0].includes('JB CREATION');
                if (aIsJB && !bIsJB) return -1;
                if (!aIsJB && bIsJB) return 1;
                return a[0].localeCompare(b[0]);
            });
            
            flipkartSellers.forEach(([seller, count]) => {
                breakdownHtml += `
                    <div class="platform-breakdown-item platform-breakdown-sub">
                        <div class="platform-breakdown-icon">↳</div>
                        <span class="platform-breakdown-name">${seller}</span>
                        <span class="platform-breakdown-count">${count}</span>
                        <span class="platform-breakdown-label">labels</span>
                    </div>
                `;
            });
        }
        breakdownHtml += `</div>`; // Close flipkart-section
    }
    
    // AMAZON Section
    if (platformCounts.amazon > 0) {
        breakdownHtml += `<div class="platform-section amazon-section">`;
        breakdownHtml += `
            <div class="platform-breakdown-item platform-breakdown-main">
                <div class="platform-breakdown-icon">📦</div>
                <span class="platform-breakdown-name">Amazon</span>
                <span class="platform-breakdown-count">${platformCounts.amazon}</span>
                <span class="platform-breakdown-label">labels processed</span>
            </div>
        `;
        breakdownHtml += `</div>`; // Close amazon-section
    }
    
    // UNKNOWN Section
    if (platformCounts.unknown > 0) {
        breakdownHtml += `<div class="platform-section unknown-section">`;
        breakdownHtml += `
            <div class="platform-breakdown-item platform-breakdown-main">
                <div class="platform-breakdown-icon">❓</div>
                <span class="platform-breakdown-name">Unknown</span>
                <span class="platform-breakdown-count">${platformCounts.unknown}</span>
                <span class="platform-breakdown-label">labels processed</span>
            </div>
        `;
        breakdownHtml += `</div>`; // Close unknown-section
    }
    
    platformBreakdownStats.innerHTML = breakdownHtml;
    
    // Display matched labels WITH COUNTS (sorted by count, high to low)
    const labelChips = document.getElementById('labelChips');
    // Sort labels by count in descending order
    const sortedLabels = labels.sort((a, b) => (labelCounts[b] || 0) - (labelCounts[a] || 0));
    labelChips.innerHTML = sortedLabels.map(label => {
        const count = labelCounts[label] || 0;
        return `<span class="label-chip">${label} <span class="label-count">×${count}</span></span>`;
    }).join('');
    
    // Show results section
    resultsSection.style.display = 'block';
    
    // Scroll to results
    resultsSection.scrollIntoView({ behavior: 'smooth' });
}

async function downloadSortedPDF() {
    if (!processedPDF) return;
    
    try {
        const pdfBytes = await processedPDF.save();
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = `sorted_labels_${new Date().getTime()}.pdf`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
    } catch (error) {
        console.error('Error downloading PDF:', error);
        alert('Error downloading PDF. Please try again.');
    }
}

function updateProgress(percent, message) {
    progressFill.style.width = `${percent}%`;
    progressFill.textContent = `${Math.round(percent)}%`;
    statusText.textContent = message;
}

// Password and Edit Modal Functions
function showPasswordModal(action) {
    pendingPasswordAction = action;
    passwordModal.style.display = 'flex';
    passwordInput.value = '';
    passwordError.style.display = 'none';
    
    // Update the message based on action
    const messageEl = document.getElementById('passwordModalMessage');
    if (messageEl) {
        if (action === 'editLabels') {
            messageEl.textContent = 'Enter password to edit priority labels';
        } else if (action === 'googleSheets') {
            messageEl.textContent = 'Enter password to configure Google Sheets';
        } else if (action === 'labelCriteria') {
            messageEl.textContent = 'Enter password to configure label detection settings';
        } else if (action === 'finderDelete') {
            messageEl.textContent = 'Enter password to delete this history file';
        }
    }
    
    passwordInput.focus();
}

function hidePasswordModal() {
    passwordModal.style.display = 'none';
    passwordInput.value = '';
    passwordError.style.display = 'none';
}

async function checkPassword() {
    const enteredPassword = passwordInput.value;
    
    if (enteredPassword === EDIT_PASSWORD) {
        hidePasswordModal();
        // Execute the pending action based on what was requested
        if (pendingPasswordAction === 'editLabels') {
            showEditLabelsModal();
        } else if (pendingPasswordAction === 'googleSheets') {
            showGoogleSheetsModal();
        } else if (pendingPasswordAction === 'labelCriteria') {
            showLabelCriteriaModal();
        } else if (pendingPasswordAction === 'finderDelete') {
            await confirmFinderHistoryDelete();
        }
        pendingPasswordAction = null;
    } else {
        passwordError.style.display = 'block';
        passwordInput.value = '';
        passwordInput.focus();
    }
}

function showEditLabelsModal() {
    editLabelsModal.style.display = 'flex';
    labelsTextarea.value = PRIORITY_LABELS.join('\n');
    labelsTextarea.focus();
}

function hideEditLabelsModal() {
    editLabelsModal.style.display = 'none';
}

async function saveLabels() {
    const newLabels = labelsTextarea.value
        .split('\n')
        .map(label => label.trim())
        .filter(label => label.length > 0);
    
    if (newLabels.length === 0) {
        alert('Please enter at least one label!');
        return;
    }
    
    // Check if Google Sheets is configured
    if (!GOOGLE_SHEETS_CONFIG.webAppUrl) {
        alert('⚠️ Google Sheets is not configured!\n\nPlease configure Google Sheets integration first to save priority labels globally.');
        return;
    }
    
    // Show saving indicator
    saveLabelsBtn.disabled = true;
    saveLabelsBtn.textContent = '⏳ Saving...';
    
    try {
        // Save to Google Sheets
        const response = await fetch(GOOGLE_SHEETS_CONFIG.webAppUrl, {
            method: 'POST',
            mode: 'no-cors',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                action: 'savePriorityLabels',
                labels: newLabels
            })
        });
        
        // Update local state
        PRIORITY_LABELS = newLabels;
        displayPriorityList();
        hideEditLabelsModal();
        
        // Show success message
        alert(`✅ Priority labels saved globally!\n\nTotal labels: ${PRIORITY_LABELS.length}\n\nAll users will now see these updated labels.`);
        
    } catch (error) {
        console.error('Error saving labels to Google Sheets:', error);
        alert(`❌ Error saving labels: ${error.message}\n\nLabels were updated locally but may not persist.`);
        
        // Still update locally even if save fails
        PRIORITY_LABELS = newLabels;
        displayPriorityList();
        hideEditLabelsModal();
    } finally {
        saveLabelsBtn.disabled = false;
        saveLabelsBtn.textContent = '💾 Save Labels';
    }
}

// Cropping functions for different platforms (to be implemented)
const croppingLogic = {
    flipkart: {
        // Placeholder for Flipkart-specific cropping logic
        cropPage: async (page) => {
            // You can add custom cropping dimensions here
            console.log('Applying Flipkart cropping logic');
            return page;
        }
    },
    amazon: {
        // Placeholder for Amazon-specific cropping logic
        cropPage: async (page) => {
            console.log('Applying Amazon cropping logic');
            return page;
        }
    },
    meesho: {
        // Placeholder for Meesho-specific cropping logic
        cropPage: async (page) => {
            console.log('Applying Meesho cropping logic');
            return page;
        }
    }
};

// Export for future use
window.croppingLogic = croppingLogic;

// ============================================
// GOOGLE SHEETS INTEGRATION
// ============================================

// DOM elements for Google Sheets (will be initialized after DOM loads)
let googleSheetsModal, sheetsIdInput, sheetNameInput, webAppUrlInput;
let saveSheetsConfigBtn, cancelSheetsConfigBtn, configureGoogleSheetsBtn;
let stockUpdateStatus;

// Initialize Google Sheets elements after DOM is loaded
function initializeGoogleSheetsElements() {
    googleSheetsModal = document.getElementById('googleSheetsModal');
    sheetsIdInput = document.getElementById('sheetsIdInput');
    sheetNameInput = document.getElementById('sheetNameInput');
    webAppUrlInput = document.getElementById('webAppUrlInput');
    saveSheetsConfigBtn = document.getElementById('saveSheetsConfigBtn');
    cancelSheetsConfigBtn = document.getElementById('cancelSheetsConfigBtn');
    configureGoogleSheetsBtn = document.getElementById('configureGoogleSheetsBtn');
    stockUpdateStatus = document.getElementById('stockUpdateStatus');
    
    // Setup Google Sheets event listeners
    if (configureGoogleSheetsBtn) {
        configureGoogleSheetsBtn.addEventListener('click', () => showPasswordModal('googleSheets'));
    }
    if (saveSheetsConfigBtn) {
        saveSheetsConfigBtn.addEventListener('click', saveGoogleSheetsConfig);
    }
    if (cancelSheetsConfigBtn) {
        cancelSheetsConfigBtn.addEventListener('click', hideGoogleSheetsModal);
    }
    if (googleSheetsModal) {
        googleSheetsModal.addEventListener('click', (e) => {
            if (e.target === googleSheetsModal) {
                hideGoogleSheetsModal();
            }
        });
    }
}

// Call this from initializeApp
const originalInitializeApp = initializeApp;
initializeApp = function() {
    originalInitializeApp();
    initializeGoogleSheetsElements();
    updateGoogleSheetsStatus();
};

function showGoogleSheetsModal() {
    if (!googleSheetsModal) return;
    googleSheetsModal.style.display = 'flex';
    if (sheetsIdInput) sheetsIdInput.value = GOOGLE_SHEETS_CONFIG.spreadsheetId;
    if (sheetNameInput) sheetNameInput.value = GOOGLE_SHEETS_CONFIG.sheetName;
    if (webAppUrlInput) webAppUrlInput.value = GOOGLE_SHEETS_CONFIG.webAppUrl;
}

function hideGoogleSheetsModal() {
    if (!googleSheetsModal) return;
    googleSheetsModal.style.display = 'none';
}

function saveGoogleSheetsConfig() {
    GOOGLE_SHEETS_CONFIG.spreadsheetId = sheetsIdInput.value.trim();
    GOOGLE_SHEETS_CONFIG.sheetName = sheetNameInput.value.trim() || 'STOCK COUNT';
    GOOGLE_SHEETS_CONFIG.webAppUrl = webAppUrlInput.value.trim();
    
    // Save to localStorage
    localStorage.setItem('googleSheetsId', GOOGLE_SHEETS_CONFIG.spreadsheetId);
    localStorage.setItem('googleSheetName', GOOGLE_SHEETS_CONFIG.sheetName);
    localStorage.setItem('googleWebAppUrl', GOOGLE_SHEETS_CONFIG.webAppUrl);
    
    // Save to cloud (async, non-blocking)
    saveAppConfigToCloud(GOOGLE_SHEETS_CONFIG);
    
    hideGoogleSheetsModal();
    updateGoogleSheetsStatus();
    alert('✅ Google Sheets configuration saved and syncing to cloud!');
}

function updateGoogleSheetsStatus() {
    const statusEl = document.getElementById('googleSheetsConnectionStatus');
    if (!statusEl) return;
    
    if (GOOGLE_SHEETS_CONFIG.webAppUrl) {
        statusEl.innerHTML = '<span class="status-connected">✅ Connected</span>';
    } else {
        statusEl.innerHTML = '<span class="status-disconnected">⚠️ Not Configured</span>';
    }
}

// Function to deduct stock from Google Sheets
async function deductStockFromGoogleSheets(labelCounts) {
    if (!GOOGLE_SHEETS_CONFIG.webAppUrl) {
        console.log('Google Sheets not configured, skipping stock update');
        return { success: false, message: 'Google Sheets not configured' };
    }
    
    if (Object.keys(labelCounts).length === 0) {
        return { success: false, message: 'No labels to update' };
    }
    
    try {
        // Show updating status
        if (stockUpdateStatus) {
            stockUpdateStatus.innerHTML = '<span class="status-updating">⏳ Updating stock...</span>';
            stockUpdateStatus.style.display = 'block';
        }
        
        // Prepare the data to send
        const updateData = {
            action: 'deductStock',
            sheetName: GOOGLE_SHEETS_CONFIG.sheetName,
            labelCounts: labelCounts
        };
        
        // Use POST with redirect:follow for Google Apps Script
        // Google Apps Script returns a redirect, so we need to follow it
        const response = await fetch(GOOGLE_SHEETS_CONFIG.webAppUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'text/plain', // Use text/plain to avoid CORS preflight
            },
            body: JSON.stringify(updateData),
            redirect: 'follow'
        });
        
        // Try to parse the response
        let result;
        try {
            const text = await response.text();
            result = JSON.parse(text);
        } catch (parseError) {
            // If parsing fails, check if response was ok
            if (response.ok) {
                result = { success: true };
            } else {
                throw new Error('Failed to update stock: ' + response.status);
            }
        }
        
        if (result.success) {
            if (stockUpdateStatus) {
                const updatedCount = result.totalUpdated || Object.keys(labelCounts).length;
                stockUpdateStatus.innerHTML = `<span class="status-success">✅ Stock updated! (${updatedCount} labels deducted)</span>`;
            }
            console.log('Stock deduction successful:', result);
            return { success: true, message: 'Stock updated successfully', data: result };
        } else {
            throw new Error(result.message || 'Unknown error');
        }
        
    } catch (error) {
        console.error('Error updating Google Sheets:', error);
        if (stockUpdateStatus) {
            stockUpdateStatus.innerHTML = `<span class="status-error">❌ Error: ${error.message}</span>`;
        }
        return { success: false, message: error.message };
    }
}

// Updated download function with stock deduction
async function downloadSortedPDF() {
    if (!processedPDF) return;
    
    try {
        const pdfBytes = await processedPDF.save();
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = `sorted_labels_${new Date().getTime()}.pdf`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
    } catch (error) {
        console.error('Error downloading PDF:', error);
        alert('Error downloading PDF. Please try again.');
    }
}

async function updateStockToGoogleSheets() {
    if (!processedPDF) {
        alert('Please process labels first before updating stock.');
        return;
    }

    if (!GOOGLE_SHEETS_CONFIG.webAppUrl) {
        alert('⚠️ Google Sheets is not configured. Please configure it first.');
        return;
    }

    if (!labelOccurrences || Object.keys(labelOccurrences).length === 0) {
        alert('No matched labels found to update stock.');
        return;
    }

    const canCaptureForCurrentBatch = Boolean(currentProcessedBatchToken) && lastFinderCapturedBatchToken !== currentProcessedBatchToken;

    if (stockAlreadyDeducted) {
        if (!finderCapturedForBatch && canCaptureForCurrentBatch) {
            const recovered = await captureFinderEntryFromCurrentBatch();
            finderCapturedForBatch = recovered;
            if (recovered) {
                lastFinderCapturedBatchToken = currentProcessedBatchToken;
            }
        }
        if (stockUpdateStatus) {
            stockUpdateStatus.innerHTML = '<span class="status-success">✅ Stock already updated for this batch</span>';
            stockUpdateStatus.style.display = 'block';
        }
        return;
    }

    const result = await deductStockFromGoogleSheets(labelOccurrences);
    if (result.success) {
        stockAlreadyDeducted = true;
        if (canCaptureForCurrentBatch) {
            finderCapturedForBatch = await captureFinderEntryFromCurrentBatch();
            if (finderCapturedForBatch) {
                lastFinderCapturedBatchToken = currentProcessedBatchToken;
            }
        }
    }
}

function formatFinderTimestamp(dateInput) {
    const date = dateInput instanceof Date ? dateInput : new Date(dateInput);
    return date.toLocaleString('en-IN', {
        day: '2-digit',
        month: 'short',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: true
    });
}

function normalizeFinderText(text) {
    return String(text || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function getFinderEntryMeta(entry) {
    return {
        id: entry.id,
        fileName: entry.fileName,
        createdAt: entry.createdAt,
        createdAtDisplay: entry.createdAtDisplay,
        pageCount: entry.pageCount || 0,
        trackingCount: Array.isArray(entry.trackings) ? entry.trackings.length : 0,
        trackings: Array.isArray(entry.trackings) ? entry.trackings : []
    };
}

function openFinderDb() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(LABEL_FINDER_DB, 1);

        request.onupgradeneeded = (event) => {
            const db = event.target.result;
            if (!db.objectStoreNames.contains(LABEL_FINDER_STORE)) {
                db.createObjectStore(LABEL_FINDER_STORE, { keyPath: 'id' });
            }
        };

        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error || new Error('Failed to open finder DB'));
    });
}

function idbRequestToPromise(request) {
    return new Promise((resolve, reject) => {
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error || new Error('IndexedDB request failed'));
    });
}

function isFinderQuotaError(error) {
    const errorText = String((error && error.message) || error || '').toLowerCase();
    return (error && error.name === 'QuotaExceededError') || errorText.includes('quota') || errorText.includes('space');
}

function getFinderPdfCacheKey(entryId) {
    return `${window.location.origin}/__finder_pdf__/${encodeURIComponent(entryId)}.pdf`;
}

async function putFinderPdfInCache(entryId, pdfBlob) {
    if (!FINDER_CAN_USE_CACHE_API) return;
    const cache = await caches.open(LABEL_FINDER_CACHE);
    await cache.put(
        getFinderPdfCacheKey(entryId),
        new Response(pdfBlob, {
            headers: {
                'Content-Type': 'application/pdf'
            }
        })
    );
}

async function getFinderPdfFromCache(entryId) {
    if (!FINDER_CAN_USE_CACHE_API) return null;
    const cache = await caches.open(LABEL_FINDER_CACHE);
    const response = await cache.match(getFinderPdfCacheKey(entryId));
    if (!response) return null;
    return await response.blob();
}

async function deleteFinderPdfFromCache(entryId) {
    if (!FINDER_CAN_USE_CACHE_API) return;
    const cache = await caches.open(LABEL_FINDER_CACHE);
    await cache.delete(getFinderPdfCacheKey(entryId));
}

async function saveFinderEntryAndPdfWithEviction(entry, pdfBlob) {
    let lastError = null;

    for (let attempt = 0; attempt <= LABEL_FINDER_LIMIT; attempt++) {
        try {
            if (FINDER_CAN_USE_CACHE_API) {
                await putFinderPdfInCache(entry.id, pdfBlob);
                await saveFinderEntryWithEviction(entry);
            } else {
                await saveFinderEntryWithEviction({ ...entry, pdfBlob });
            }
            return;
        } catch (error) {
            lastError = error;
            if (!isFinderQuotaError(error)) {
                throw error;
            }

            const existing = await getAllFinderEntries();
            if (!existing.length) {
                throw error;
            }

            const oldest = existing
                .sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt))[0];

            if (!oldest || !oldest.id) {
                throw error;
            }

            await deleteFinderEntry(oldest.id);
            finderPdfCache.delete(oldest.id);
        }
    }

    throw lastError || new Error('Unable to save finder entry and PDF');
}

async function saveFinderEntry(entry) {
    const db = await openFinderDb();
    try {
        const tx = db.transaction(LABEL_FINDER_STORE, 'readwrite');
        const store = tx.objectStore(LABEL_FINDER_STORE);
        await idbRequestToPromise(store.put(entry));
        await new Promise((resolve, reject) => {
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error || new Error('Failed saving finder entry'));
            tx.onabort = () => reject(tx.error || new Error('Finder entry save aborted'));
        });
    } finally {
        db.close();
    }
}

async function saveFinderEntryWithEviction(entry) {
    let lastError = null;

    for (let attempt = 0; attempt <= LABEL_FINDER_LIMIT; attempt++) {
        try {
            await saveFinderEntry(entry);
            return;
        } catch (error) {
            lastError = error;
            const isQuotaIssue = isFinderQuotaError(error);

            if (!isQuotaIssue) {
                throw error;
            }

            const existing = await getAllFinderEntries();
            if (!existing.length) {
                throw error;
            }

            const oldest = existing
                .sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt))[0];

            if (!oldest || !oldest.id) {
                throw error;
            }

            await deleteFinderEntry(oldest.id);
            finderPdfCache.delete(oldest.id);
        }
    }

    throw lastError || new Error('Unable to save finder entry');
}

async function getAllFinderEntries() {
    const db = await openFinderDb();
    try {
        const tx = db.transaction(LABEL_FINDER_STORE, 'readonly');
        const store = tx.objectStore(LABEL_FINDER_STORE);
        const result = await idbRequestToPromise(store.getAll());
        return Array.isArray(result) ? result : [];
    } finally {
        db.close();
    }
}

async function deleteFinderEntry(entryId) {
    await deleteFinderPdfFromCache(entryId);
    const db = await openFinderDb();
    try {
        const tx = db.transaction(LABEL_FINDER_STORE, 'readwrite');
        const store = tx.objectStore(LABEL_FINDER_STORE);
        await idbRequestToPromise(store.delete(entryId));
        await new Promise((resolve, reject) => {
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error || new Error('Failed deleting finder entry'));
            tx.onabort = () => reject(tx.error || new Error('Finder entry delete aborted'));
        });
    } finally {
        db.close();
    }
}

async function refreshFinderFromStorage() {
    const entries = await getAllFinderEntries();
    const sortedEntries = entries.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

    finderPdfCache.clear();
    finderEntries = sortedEntries.map(entry => {
        if (entry.pdfBlob) {
            finderPdfCache.set(entry.id, entry.pdfBlob);
        }
        return getFinderEntryMeta(entry);
    });

    buildFinderIndex();
    renderFinderHistory();
}

async function confirmFinderHistoryDelete() {
    if (!pendingFinderDeleteEntryId) return;

    const entryId = pendingFinderDeleteEntryId;
    pendingFinderDeleteEntryId = null;

    await deleteFinderEntry(entryId);
    await refreshFinderFromStorage();

    if (finderCurrentSelection && finderCurrentSelection.entryId === entryId) {
        resetFinderSearchStateAfterPrint();
    }
}

function buildFinderIndex() {
    finderIndex = [];

    finderEntries.forEach(entry => {
        const timestampMs = new Date(entry.createdAt).getTime();
        entry.trackings.forEach(trackingInfo => {
            const normalized = normalizeFinderText(trackingInfo.tracking);
            if (!normalized) return;

            finderIndex.push({
                entryId: entry.id,
                fileName: entry.fileName,
                createdAt: entry.createdAt,
                createdAtDisplay: entry.createdAtDisplay,
                timestampMs,
                tracking: trackingInfo.tracking,
                trackingNormalized: normalized,
                labelPageIndex: trackingInfo.labelPageIndex,
                invoicePageIndex: typeof trackingInfo.invoicePageIndex === 'number' ? trackingInfo.invoicePageIndex : null
            });
        });
    });

    finderIndex.sort((a, b) => b.timestampMs - a.timestampMs);
}

function renderFinderHistory() {
    if (!finderHistoryList) return;

    if (finderEntries.length === 0) {
        finderHistoryList.innerHTML = '';
        return;
    }

    finderHistoryList.innerHTML = finderEntries
        .map(entry => `
            <div class="finder-history-item">
                <div class="finder-item-meta">
                    <div class="finder-item-title">${entry.fileName}</div>
                    <div class="finder-item-subtitle">${entry.createdAtDisplay} · ${entry.pageCount} pages · ${entry.trackingCount} tracking numbers</div>
                </div>
                <div class="finder-history-actions">
                    <button class="btn btn-secondary" data-finder-entry-open="${entry.id}">Open</button>
                    <button class="btn btn-secondary" data-finder-entry-delete="${entry.id}">Delete</button>
                </div>
            </div>
        `)
        .join('');

    finderHistoryList.querySelectorAll('[data-finder-entry-open]').forEach(button => {
        button.addEventListener('click', async event => {
            event.preventDefault();
            const entryId = button.getAttribute('data-finder-entry-open');
            if (!entryId) return;
            await openFinderEntryPdf(entryId);
        });
    });

    finderHistoryList.querySelectorAll('[data-finder-entry-delete]').forEach(button => {
        button.addEventListener('click', event => {
            event.preventDefault();
            const entryId = button.getAttribute('data-finder-entry-delete');
            if (!entryId) return;
            requestFinderHistoryDelete(entryId);
        });
    });
}

function renderFinderResults(matches, query) {
    if (!finderResultsList) return;

    if (!query || query.length < 3) {
        finderRenderedMatches = [];
        finderResultsList.innerHTML = '';
        return;
    }

    if (!matches.length) {
        finderRenderedMatches = [];
        finderResultsList.innerHTML = '';
        return;
    }

    finderRenderedMatches = matches.slice();

    finderResultsList.innerHTML = matches
        .map((match, index) => `
            <div class="finder-result-item">
                <div class="finder-item-meta">
                    <div class="finder-item-title">${match.tracking}</div>
                    <div class="finder-item-subtitle">${match.fileName} · ${match.createdAtDisplay}</div>
                </div>
                <button class="btn btn-secondary" data-finder-match-index="${index}">Open</button>
            </div>
        `)
        .join('');

    finderResultsList.querySelectorAll('[data-finder-match-index]').forEach(button => {
        button.addEventListener('click', async event => {
            event.preventDefault();
            const index = Number(button.getAttribute('data-finder-match-index'));
            const match = finderRenderedMatches[index];
            if (!match) return;
            await openFinderMatchPdf(match);
        });
    });
}

async function initializeLabelFinder() {
    if (!finderHistoryList) return;

    try {
        await refreshFinderFromStorage();
        renderFinderResults([], '');
    } catch (error) {
        console.error('Failed to initialize label finder:', error);
        if (finderSearchHelp) {
            finderSearchHelp.textContent = '';
        }
    }
}

function extractTrackingCandidates(pageText) {
    const text = String(pageText || '');
    const regexes = [
        /\b\d{12,18}\b/g,
        /\b[A-Z]{2}\d{10,16}[A-Z]{0,6}\b/g,
        /\b[A-Z]{3,8}\d{7,14}[A-Z]{0,6}\b/g,
        /\b\d{8,14}[A-Z]{2,6}\b/g
    ];

    const candidates = new Set();
    const normalizedText = text.toUpperCase();

    regexes.forEach(regex => {
        const matches = normalizedText.match(regex) || [];
        matches.forEach(match => {
            const cleaned = normalizeFinderText(match);
            if (cleaned.length >= 10 && cleaned.length <= 24) {
                candidates.add(cleaned);
            }
        });
    });

    return Array.from(candidates);
}

async function extractTrackingsFromSortedPdf(pdfBytes) {
    const pdfDoc = await pdfjsLib.getDocument({ data: pdfBytes.slice(0) }).promise;
    const pageCount = pdfDoc.numPages;
    const trackingRecords = [];
    const seen = new Set();

    for (let pageNumber = 1; pageNumber <= pageCount; pageNumber++) {
        const page = await pdfDoc.getPage(pageNumber);
        const textContent = await page.getTextContent();
        const pageText = textContent.items.map(item => item.str || '').join(' ');
        const hasLabelSignal = hasBarcodeLikeTextForSorter(String(pageText || '').toLowerCase()) || /AWB|TRACKING|SHIP(?:PING)?|PICKUP|DELIVERY|RETURN\s+TO/i.test(pageText);
        if (!hasLabelSignal) continue;

        const candidates = extractTrackingCandidates(pageText);
        const labelPageIndex = pageNumber - 1;
        const invoicePageIndex = pageNumber < pageCount ? pageNumber : null;

        candidates.forEach(tracking => {
            const key = `${tracking}::${labelPageIndex}`;
            if (seen.has(key)) return;
            seen.add(key);
            trackingRecords.push({ tracking, labelPageIndex, invoicePageIndex });
        });
    }

    return { pageCount, trackingRecords };
}

async function captureFinderEntryFromCurrentBatch() {
    if (!processedPDF) return false;

    try {
        const pdfBytes = await processedPDF.save();
        const pdfBlob = new Blob([pdfBytes], { type: 'application/pdf' });
        const { pageCount, trackingRecords } = await extractTrackingsFromSortedPdf(pdfBytes);

        const createdAt = new Date().toISOString();
        const entry = {
            id: `finder_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`,
            createdAt,
            createdAtDisplay: formatFinderTimestamp(createdAt),
            fileName: latestSortedPdfName || `sorted_labels_${Date.now()}.pdf`,
            pageCount,
            trackings: trackingRecords
        };

        await saveFinderEntryAndPdfWithEviction(entry, pdfBlob);

        // Reload to keep state consistent with DB and enforce retention
        const entries = await getAllFinderEntries();
        const sorted = entries.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
        const keep = sorted.slice(0, LABEL_FINDER_LIMIT);
        const remove = sorted.slice(LABEL_FINDER_LIMIT);

        for (const oldEntry of remove) {
            await deleteFinderEntry(oldEntry.id);
            finderPdfCache.delete(oldEntry.id);
        }

        finderPdfCache.clear();
        finderEntries = keep.map(item => {
            if (item.pdfBlob) {
                finderPdfCache.set(item.id, item.pdfBlob);
            }
            return getFinderEntryMeta(item);
        });

        buildFinderIndex();
        renderFinderHistory();

        if (finderSearchHelp) {
            finderSearchHelp.textContent = '';
        }
        return true;
    } catch (error) {
        console.error('Failed to capture finder entry:', error);
        if (finderSearchHelp) {
            finderSearchHelp.textContent = '';
        }
        alert(`Label Finder history save failed: ${error.message}`);
        return false;
    }
}

function getFinderEntryById(entryId) {
    return finderEntries.find(entry => entry.id === entryId) || null;
}

async function getFinderEntryBlob(entryId) {
    if (finderPdfCache.has(entryId)) {
        return finderPdfCache.get(entryId);
    }

    const cachedBlob = await getFinderPdfFromCache(entryId);
    if (cachedBlob) {
        finderPdfCache.set(entryId, cachedBlob);
        return cachedBlob;
    }

    const db = await openFinderDb();
    try {
        const tx = db.transaction(LABEL_FINDER_STORE, 'readonly');
        const store = tx.objectStore(LABEL_FINDER_STORE);
        const entry = await idbRequestToPromise(store.get(entryId));
        if (entry && entry.pdfBlob) {
            finderPdfCache.set(entryId, entry.pdfBlob);
            return entry.pdfBlob;
        }
    } finally {
        db.close();
    }

    return null;
}

async function renderPdfPageToCanvas(pdfBlob, pageIndex, canvasEl) {
    if (!canvasEl || !pdfBlob || typeof pageIndex !== 'number') return;

    const bytes = await pdfBlob.arrayBuffer();
    const pdfDoc = await pdfjsLib.getDocument({ data: new Uint8Array(bytes) }).promise;
    const page = await pdfDoc.getPage(pageIndex + 1);
    const viewport = page.getViewport({ scale: 1.35 });
    const ctx = canvasEl.getContext('2d');

    canvasEl.width = viewport.width;
    canvasEl.height = viewport.height;

    await page.render({ canvasContext: ctx, viewport }).promise;
}

function clearFinderCanvas(canvasEl) {
    if (!canvasEl) return;
    const ctx = canvasEl.getContext('2d');
    ctx.clearRect(0, 0, canvasEl.width || 1, canvasEl.height || 1);
    canvasEl.width = 1;
    canvasEl.height = 1;
}

function resetFinderSearchStateAfterPrint() {
    if (finderSearchInput) {
        finderSearchInput.value = '';
    }
    finderLastMatches = [];
    finderCurrentSelection = null;
    if (finderPrintBtn) {
        finderPrintBtn.disabled = true;
    }
    if (finderSelectedMeta) {
        finderSelectedMeta.textContent = '';
    }
    if (finderResultsList) {
        finderResultsList.innerHTML = '';
    }
    clearFinderCanvas(finderLabelCanvas);
}

function isFinderTabActive() {
    const finderTab = document.getElementById('finder-tab-content');
    return Boolean(finderTab && finderTab.classList.contains('active'));
}

function handleFinderGlobalKeydown(event) {
    if (!isFinderTabActive() || !finderSearchInput) return;
    if (event.ctrlKey || event.metaKey || event.altKey) return;

    const target = event.target;
    const isTypingInField = target && (
        target.tagName === 'INPUT' ||
        target.tagName === 'TEXTAREA' ||
        target.isContentEditable
    );

    if (event.key === 'Enter') {
        if (finderPrintBtn && !finderPrintBtn.disabled) {
            event.preventDefault();
            printFinderSelection();
        }
        return;
    }

    if (isTypingInField) return;

    if (event.key === 'Backspace') {
        event.preventDefault();
        finderSearchInput.focus();
        finderSearchInput.value = finderSearchInput.value.slice(0, -1);
        finderSearchInput.dispatchEvent(new Event('input', { bubbles: true }));
        return;
    }

    if (event.key.length === 1 && /[0-9A-Za-z]/.test(event.key)) {
        event.preventDefault();
        finderSearchInput.focus();
        finderSearchInput.value += event.key;
        finderSearchInput.dispatchEvent(new Event('input', { bubbles: true }));
    }
}

function getOrCreateFinderPrintFrame() {
    if (finderPrintFrame && document.body.contains(finderPrintFrame)) {
        return finderPrintFrame;
    }

    finderPrintFrame = document.createElement('iframe');
    finderPrintFrame.id = 'finderPrintFrame';
    finderPrintFrame.style.position = 'fixed';
    finderPrintFrame.style.right = '0';
    finderPrintFrame.style.bottom = '0';
    finderPrintFrame.style.width = '0';
    finderPrintFrame.style.height = '0';
    finderPrintFrame.style.border = '0';
    finderPrintFrame.style.opacity = '0';
    finderPrintFrame.setAttribute('aria-hidden', 'true');
    document.body.appendChild(finderPrintFrame);
    return finderPrintFrame;
}

async function selectFinderMatch(match) {
    if (!match) return;

    finderCurrentSelection = match;

    const entry = getFinderEntryById(match.entryId);
    const blob = await getFinderEntryBlob(match.entryId);

    if (!entry || !blob) {
        alert('Unable to load selected PDF entry.');
        return;
    }

    await renderPdfPageToCanvas(blob, match.labelPageIndex, finderLabelCanvas);

    finderPrintBtn.disabled = false;

    finderSelectedMeta.textContent = `Tracking: ${match.tracking} · ${entry.fileName} · ${entry.createdAtDisplay}`;
}

async function createFinderSubsetPdfBlob(sourceBlob, pageIndexes) {
    const { PDFDocument } = PDFLib;
    const sourceBytes = await sourceBlob.arrayBuffer();
    const sourceDoc = await PDFDocument.load(sourceBytes);
    const outputDoc = await PDFDocument.create();
    const pages = await outputDoc.copyPages(sourceDoc, pageIndexes);
    pages.forEach(page => outputDoc.addPage(page));
    const outputBytes = await outputDoc.save();
    return new Blob([outputBytes], { type: 'application/pdf' });
}

function openBlobInBackgroundTab(blob) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.target = '_blank';
    anchor.rel = 'noopener noreferrer';
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);

    try {
        window.focus();
    } catch (error) {
        console.warn('Unable to preserve current tab focus after opening PDF tab:', error);
    }

    setTimeout(() => {
        try {
            URL.revokeObjectURL(url);
        } catch (error) {
            console.warn('Failed to revoke finder PDF URL:', error);
        }
    }, 120000);
}

let finderLastOpenRequestKey = '';
let finderLastOpenRequestAt = 0;

function shouldSkipDuplicateFinderOpen(requestKey) {
    const now = Date.now();
    if (finderLastOpenRequestKey === requestKey && (now - finderLastOpenRequestAt) < 700) {
        return true;
    }

    finderLastOpenRequestKey = requestKey;
    finderLastOpenRequestAt = now;
    return false;
}

async function openFinderMatchPdf(match) {
    if (!match) return;

    const requestKey = `match:${match.entryId}:${match.labelPageIndex}:${match.tracking}`;
    if (shouldSkipDuplicateFinderOpen(requestKey)) return;

    const sourceBlob = await getFinderEntryBlob(match.entryId);
    if (!sourceBlob) {
        alert('Unable to load selected PDF entry.');
        return;
    }

    const labelBlob = await createFinderSubsetPdfBlob(sourceBlob, [match.labelPageIndex]);
    openBlobInBackgroundTab(labelBlob);
}

async function openFinderEntryPdf(entryId) {
    const requestKey = `entry:${entryId}`;
    if (shouldSkipDuplicateFinderOpen(requestKey)) return;

    const blob = await getFinderEntryBlob(entryId);
    if (!blob) {
        alert('Unable to load selected PDF entry.');
        return;
    }

    openBlobInBackgroundTab(blob);
}

function getFinderMatches(queryRaw) {
    const query = normalizeFinderText(queryRaw);
    if (!query || query.length < 3) return [];

    const unique = new Set();
    const matches = [];

    for (const row of finderIndex) {
        if (!row.trackingNormalized.includes(query)) continue;

        const key = `${row.entryId}::${row.labelPageIndex}::${row.tracking}`;
        if (unique.has(key)) continue;
        unique.add(key);
        matches.push(row);
    }

    return matches;
}

let finderLastMatches = [];
let finderRenderedMatches = [];

function isSameFinderMatch(a, b) {
    if (!a || !b) return false;

    return a.entryId === b.entryId &&
        a.labelPageIndex === b.labelPageIndex &&
        a.tracking === b.tracking;
}

function handleFinderSearchInput() {
    clearTimeout(finderSearchDebounceTimer);

    finderSearchDebounceTimer = setTimeout(async () => {
        const query = finderSearchInput.value.trim();
        const matches = getFinderMatches(query);
        finderLastMatches = matches;

        renderFinderResults(matches, normalizeFinderText(query));

        if (matches.length === 1) {
            await selectFinderMatch(matches[0]);
        } else {
            const selectedStillVisible = finderCurrentSelection
                ? matches.some(match => isSameFinderMatch(match, finderCurrentSelection))
                : false;

            if (!selectedStillVisible) {
                finderCurrentSelection = null;
                finderPrintBtn.disabled = true;
                finderSelectedMeta.textContent = '';
                clearFinderCanvas(finderLabelCanvas);
            }
        }

        if (finderSearchHelp) {
            finderSearchHelp.textContent = '';
        }
    }, 60);
}

async function printFinderSelection() {
    if (!finderCurrentSelection) {
        alert('Please select a label result first.');
        return;
    }

    const blob = await getFinderEntryBlob(finderCurrentSelection.entryId);
    if (!blob) {
        alert('Unable to load selected PDF for printing.');
        return;
    }

    try {
        const { PDFDocument } = PDFLib;
        const sourceBytes = await blob.arrayBuffer();
        const sourceDoc = await PDFDocument.load(sourceBytes);
        const printDoc = await PDFDocument.create();
        const includeInvoice = !finderIncludeInvoiceToggle || finderIncludeInvoiceToggle.checked;
        const pagesToCopy = [finderCurrentSelection.labelPageIndex];

        if (includeInvoice && typeof finderCurrentSelection.invoicePageIndex === 'number') {
            pagesToCopy.push(finderCurrentSelection.invoicePageIndex);
        }

        const copiedPages = await printDoc.copyPages(sourceDoc, pagesToCopy);
        copiedPages.forEach(page => printDoc.addPage(page));

        const outputBytes = await printDoc.save();
        const outputBlob = new Blob([outputBytes], { type: 'application/pdf' });
        const outputUrl = URL.createObjectURL(outputBlob);

        const printFrame = getOrCreateFinderPrintFrame();

        if (finderPrintUrlToRevoke) {
            try {
                URL.revokeObjectURL(finderPrintUrlToRevoke);
            } catch (revokeError) {
                console.warn('Finder print URL revoke warning:', revokeError);
            }
            finderPrintUrlToRevoke = null;
        }

        finderPrintUrlToRevoke = outputUrl;
        printFrame.onload = () => {
            try {
                printFrame.contentWindow.focus();
                printFrame.contentWindow.print();
                resetFinderSearchStateAfterPrint();
            } catch (printError) {
                console.error('Finder print frame error:', printError);
                alert(`Print failed: ${printError.message}`);
            }
        };

        printFrame.src = outputUrl;
    } catch (error) {
        console.error('Finder print failed:', error);
        alert(`Print failed: ${error.message}`);
    }
}

window.openFinderEntry = async function(entryId) {
    await openFinderEntryPdf(entryId);
};

window.openFinderMatch = async function(matchIndex) {
    const match = finderRenderedMatches[matchIndex] || finderLastMatches[matchIndex];
    if (!match) return;
    await openFinderMatchPdf(match);
};

window.requestFinderHistoryDelete = function(entryId) {
    pendingFinderDeleteEntryId = entryId;
    showPasswordModal('finderDelete');
};

// ============================================
// GOOGLE APPS SCRIPT CODE (for user to deploy)
// ============================================
/*
To make the Google Sheets integration work, you need to create a Google Apps Script:

1. Open your Google Sheet
2. Go to Extensions > Apps Script
3. Delete any existing code and paste the following:

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    
    if (data.action === 'deductStock') {
      var sheetName = data.sheetName || 'Sheet1';
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      
      if (!sheet) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          message: 'Sheet not found: ' + sheetName
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      var labelCounts = data.labelCounts;
      var dataRange = sheet.getDataRange();
      var values = dataRange.getValues();
      
      // Find and update each label
      var updates = [];
      for (var label in labelCounts) {
        var deductAmount = labelCounts[label];
        
        // Search for the label in column A (case-insensitive)
        for (var i = 1; i < values.length; i++) { // Start from 1 to skip header
          var cellLabel = String(values[i][0]).toLowerCase().trim();
          var searchLabel = label.toLowerCase().trim();
          
          if (cellLabel === searchLabel) {
            var currentStock = Number(values[i][1]) || 0;
            var newStock = Math.max(0, currentStock - deductAmount);
            sheet.getRange(i + 1, 2).setValue(newStock);
            updates.push({label: label, oldStock: currentStock, newStock: newStock, deducted: deductAmount});
            break;
          }
        }
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        message: 'Stock updated',
        updates: updates
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Unknown action'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    message: 'JB Creations Stock Update API is running'
  })).setMimeType(ContentService.MimeType.JSON);
}

4. Save the script
5. Click Deploy > New deployment
6. Select "Web app" as the type
7. Set "Execute as" to "Me"
8. Set "Who has access" to "Anyone"
9. Click Deploy and authorize
10. Copy the Web App URL and paste it in the website configuration
*/

// ============================================
// LABEL DETECTION CRITERIA SETTINGS
// ============================================

function showLabelCriteriaModal() {
    if (!labelCriteriaModal) return;
    
    // Load current criteria into form fields
    meeshoReturnAddressPattern.value = LABEL_CRITERIA.meesho.returnAddressPatternString || 
        DEFAULT_LABEL_CRITERIA.meesho.returnAddressPatternString;
    meeshoCouriers.value = (LABEL_CRITERIA.meesho.couriers || DEFAULT_LABEL_CRITERIA.meesho.couriers).join(', ');
    flipkartShippingAddressKeyword.value = LABEL_CRITERIA.flipkart.shippingAddressKeyword || 
        DEFAULT_LABEL_CRITERIA.flipkart.shippingAddressKeyword;
    flipkartSoldByPattern.value = LABEL_CRITERIA.flipkart.soldByPatternString || 
        DEFAULT_LABEL_CRITERIA.flipkart.soldByPatternString;
    
    labelCriteriaModal.style.display = 'flex';
}

function hideLabelCriteriaModal() {
    if (!labelCriteriaModal) return;
    labelCriteriaModal.style.display = 'none';
}

async function saveLabelCriteriaSettings() {
    try {
        // Validate regex patterns
        const meeshoPatternString = meeshoReturnAddressPattern.value.trim();
        const flipkartPatternString = flipkartSoldByPattern.value.trim();
        
        if (!meeshoPatternString || !flipkartPatternString) {
            alert('Please fill in all pattern fields!');
            return;
        }
        
        // Test regex patterns
        try {
            new RegExp(meeshoPatternString, 'i');
            new RegExp(flipkartPatternString, 'i');
        } catch (e) {
            alert('Invalid regex pattern! Please check your patterns.\n\nError: ' + e.message);
            return;
        }
        
        // Parse courier list
        const couriersList = meeshoCouriers.value.split(',')
            .map(c => c.trim().toUpperCase())
            .filter(c => c.length > 0);
        
        if (couriersList.length === 0) {
            alert('Please enter at least one courier name!');
            return;
        }
        
        // Show saving indicator
        saveCriteriaBtn.disabled = true;
        saveCriteriaBtn.textContent = '⏳ Saving...';
        
        // Create new criteria object
        const newCriteria = {
            meesho: {
                returnAddressPattern: new RegExp(meeshoPatternString, 'i'),
                returnAddressPatternString: meeshoPatternString,
                returnAddressFlags: 'i',
                couriers: couriersList,
                pickupKeyword: 'PICKUP'
            },
            flipkart: {
                shippingAddressKeyword: flipkartShippingAddressKeyword.value.trim(),
                soldByPattern: new RegExp(flipkartPatternString, 'i'),
                soldByPatternString: flipkartPatternString,
                soldByFlags: 'i'
            }
        };
        
        // Save the criteria
        saveLabelCriteria(newCriteria);
        
        // Reset button
        saveCriteriaBtn.disabled = false;
        saveCriteriaBtn.textContent = '💾 Save Settings';
        
        // Hide modal
        hideLabelCriteriaModal();
        
        alert('✅ Label detection criteria saved successfully!\n\nThe new criteria will be used for all future label processing.');
        
    } catch (error) {
        console.error('Error saving criteria:', error);
        alert('Error saving criteria: ' + error.message);
        saveCriteriaBtn.disabled = false;
        saveCriteriaBtn.textContent = '💾 Save Settings';
    }
}

function resetLabelCriteriaSettings() {
    if (!confirm('Are you sure you want to reset all criteria to default values?\n\nThis cannot be undone.')) {
        return;
    }
    
    const defaults = resetLabelCriteriaToDefaults();
    
    // Update form fields
    meeshoReturnAddressPattern.value = defaults.meesho.returnAddressPatternString;
    meeshoCouriers.value = defaults.meesho.couriers.join(', ');
    flipkartShippingAddressKeyword.value = defaults.flipkart.shippingAddressKeyword;
    flipkartSoldByPattern.value = defaults.flipkart.soldByPatternString;
    
    alert('✅ Criteria reset to default values!');
}


// ============================================================
// LABEL CROPPING & FORMATTING MODULE
// All code below is modular and does not alter existing logic.
// ============================================================

// --- Cropping State ---
let croppingUploadedFiles = [];
let croppingResults = []; // Array of { filename, platform, blob, pageCount, sellerId }

/**
 * Detect platform from PDF text content for cropping purposes.
 * Returns { platform: 'meesho'|'flipkart'|'amazon'|'unknown', sellerId: string }
 * This is a simplified, more robust detection than the sorting/counter logic.
 */
async function detectCroppingPlatform(pdfData) {
    const pdfjs = await pdfjsLib.getDocument({ data: pdfData.slice(0) }).promise;
    const numPages = pdfjs.numPages;

    // === AMAZON CHECK: page 2 must contain "Tax Invoice/Bill of Supply/Cash Memo" ===
    if (numPages >= 2) {
        const page2 = await pdfjs.getPage(2);
        const tc2 = await page2.getTextContent();
        const page2Text = tc2.items.map(i => i.str).join(' ');
        if (page2Text.includes('Tax Invoice/Bill of Supply/Cash Memo')) {
            return { platform: 'amazon', sellerId: 'AMAZON' };
        }
    }

    // === Read page 1 text ===
    const page1 = await pdfjs.getPage(1);
    const tc1 = await page1.getTextContent();
    const allText = tc1.items.map(i => i.str).join(' ');
    const allTextUpper = allText.toUpperCase();
    const itemStrings = tc1.items.map(i => i.str.trim()).filter(s => s);

    // === MEESHO CHECK: "return to" + (PICKUP or MEESHO or courier keywords) ===
    const hasReturnTo = /return\s+to/i.test(allText);
    const hasPickup = allTextUpper.includes('PICKUP');
    const hasMeesho = allTextUpper.includes('MEESHO');
    const hasCourier = ['DELHIVERY', 'SHADOWFAX', 'VALMO', 'XPRESS BEES', 'ECOM EXPRESS', 'EKART'].some(c => allTextUpper.includes(c));
    // Meesho labels have "return to" + a courier/PICKUP indicator
    // OR just contain "MEESHO" directly
    if (hasReturnTo && (hasPickup || hasCourier) || hasMeesho) {
        // Extract seller ID
        let sellerId = 'UNKNOWN';
        sellerId = extractMeeshoSellerId(tc1.items);
        return { platform: 'meesho', sellerId };
    }

    // === FLIPKART CHECK: "Shipping/Customer address:" or "Sold By" or "FLIPKART" ===
    const hasShippingAddr = allText.includes('Shipping/Customer address:') || allTextUpper.includes('SHIPPING/CUSTOMER ADDRESS');
    const hasSoldBy = /sold\s*by/i.test(allText);
    const hasFlipkart = allTextUpper.includes('FLIPKART') || allTextUpper.includes('EKART') || allTextUpper.includes('FK-');
    if (hasShippingAddr || (hasSoldBy && hasFlipkart)) {
        // Extract seller name
        let sellerId = 'UNKNOWN';
        const soldByMatch = allText.match(/Sold\s*[Bb]y\s*:?\s*([A-Za-z][A-Za-z0-9 _\-]+)/i);
        if (soldByMatch) {
            sellerId = soldByMatch[1].trim().replace(/\s*,.*$/, '').toUpperCase();
        }
        return { platform: 'flipkart', sellerId };
    }

    // === AMAZON fallback (single page with Amazon text) ===
    if (allTextUpper.includes('AMAZON') || allTextUpper.includes('AMAZON.IN')) {
        return { platform: 'amazon', sellerId: 'AMAZON' };
    }

    return { platform: 'unknown', sellerId: 'UNKNOWN' };
}

// --- Cropping DOM Elements (lazy-loaded) ---
function getCroppingElements() {
    return {
        uploadArea: document.getElementById('croppingUploadArea'),
        fileInput: document.getElementById('croppingFileInput'),
        clearBtn: document.getElementById('croppingClearBtn'),
        filesInfo: document.getElementById('croppingFilesInfo'),
        filesList: document.getElementById('croppingFilesList'),
        actions: document.getElementById('croppingActions'),
        autoBtn: document.getElementById('croppingAutoBtn'),
        meeshoBtn: document.getElementById('croppingMeeshoBtn'),
        flipkartBtn: document.getElementById('croppingFlipkartBtn'),
        amazonBtn: document.getElementById('croppingAmazonBtn'),
        statusSection: document.getElementById('croppingStatusSection'),
        progressFill: document.getElementById('croppingProgressFill'),
        statusText: document.getElementById('croppingStatusText'),
        resultsSection: document.getElementById('croppingResultsSection'),
        totalProcessed: document.getElementById('croppingTotalProcessed'),
        platformCount: document.getElementById('croppingPlatformCount'),
        outputList: document.getElementById('croppingOutputList'),
        downloadAllBtn: document.getElementById('croppingDownloadAllBtn'),
    };
}

// --- Init Cropping Tab ---
function initCroppingTab() {
    const el = getCroppingElements();
    if (!el.uploadArea) return;

    // Upload area click
    el.uploadArea.addEventListener('click', (e) => {
        if (e.target !== el.fileInput) el.fileInput.click();
    });

    // File input change
    el.fileInput.addEventListener('change', handleCroppingFileSelect);

    // Drag and drop
    el.uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        el.uploadArea.classList.add('dragover');
    });
    el.uploadArea.addEventListener('dragleave', () => {
        el.uploadArea.classList.remove('dragover');
    });
    el.uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        el.uploadArea.classList.remove('dragover');
        const files = Array.from(e.dataTransfer.files).filter(f => f.type === 'application/pdf');
        if (files.length > 0) handleCroppingFiles(files);
    });

    // Clear button
    el.clearBtn.addEventListener('click', clearCroppingFiles);

    // Platform buttons
    el.autoBtn.addEventListener('click', () => processCropping('automatic'));
    el.meeshoBtn.addEventListener('click', () => processCropping('meesho'));
    el.flipkartBtn.addEventListener('click', () => processCropping('flipkart'));
    el.amazonBtn.addEventListener('click', () => processCropping('amazon'));

    // Download all
    el.downloadAllBtn.addEventListener('click', downloadAllCropped);
}

function handleCroppingFileSelect(e) {
    const files = Array.from(e.target.files).filter(f => f.type === 'application/pdf');
    if (files.length > 0) handleCroppingFiles(files);
}

async function handleCroppingFiles(newFiles) {
    const el = getCroppingElements();
    // Merge with existing, deduplicate by name
    for (const file of newFiles) {
        const exists = croppingUploadedFiles.some(f => f.name === file.name && f.size === file.size);
        if (!exists) croppingUploadedFiles.push(file);
    }

    el.clearBtn.style.display = 'inline-flex';
    el.filesInfo.style.display = 'block';
    el.actions.style.display = 'block';
    el.resultsSection.style.display = 'none';

    // Show file list with page counts
    el.filesList.innerHTML = '';
    for (const file of croppingUploadedFiles) {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        const div = document.createElement('div');
        div.className = 'cropping-file-item';
        div.innerHTML = `<span class="file-name">📄 ${file.name}</span><span class="file-pages">${pdf.numPages} pages</span>`;
        el.filesList.appendChild(div);
    }
}

function clearCroppingFiles() {
    const el = getCroppingElements();
    croppingUploadedFiles = [];
    croppingResults = [];
    el.fileInput.value = '';
    el.clearBtn.style.display = 'none';
    el.filesInfo.style.display = 'none';
    el.actions.style.display = 'none';
    el.statusSection.style.display = 'none';
    el.resultsSection.style.display = 'none';
    el.filesList.innerHTML = '';
    el.outputList.innerHTML = '';
}

// --- Progress Helpers ---
function updateCroppingProgress(pct, text) {
    const el = getCroppingElements();
    el.statusSection.style.display = 'block';
    el.progressFill.style.width = `${pct}%`;
    el.statusText.textContent = text;
}

// --- Main Processing Entry Point ---
async function processCropping(mode) {
    if (croppingUploadedFiles.length === 0) {
        alert('Please upload at least one PDF file.');
        return;
    }
    const el = getCroppingElements();
    croppingResults = [];
    el.resultsSection.style.display = 'none';
    el.outputList.innerHTML = '';

    // Disable buttons during processing
    el.autoBtn.disabled = true;
    el.meeshoBtn.disabled = true;
    el.flipkartBtn.disabled = true;
    el.amazonBtn.disabled = true;

    try {
        updateCroppingProgress(5, 'Reading PDF files...');

        // Read all files into ArrayBuffers
        const fileDataArray = [];
        for (const file of croppingUploadedFiles) {
            const arrayBuffer = await file.arrayBuffer();
            fileDataArray.push({ name: file.name, data: arrayBuffer });
        }

        if (mode === 'automatic') {
            await processCroppingAutomatic(fileDataArray);
        } else {
            await processCroppingManual(fileDataArray, mode);
        }

        // Display results
        displayCroppingResults();
    } catch (err) {
        console.error('Cropping error:', err);
        updateCroppingProgress(100, `❌ Error: ${err.message}`);
    } finally {
        el.autoBtn.disabled = false;
        el.meeshoBtn.disabled = false;
        el.flipkartBtn.disabled = false;
        el.amazonBtn.disabled = false;
    }
}

// ------------------------------------------------------------------
// AUTOMATIC MODE: classify each PDF, then process per-platform
// ------------------------------------------------------------------
async function processCroppingAutomatic(fileDataArray) {
    updateCroppingProgress(10, 'Classifying labels by platform...');

    // Buckets: { meesho: [{data, pages, sellerId, name}], flipkart: [...], amazon: [...] }
    const buckets = { meesho: [], flipkart: [], amazon: [] };
    const unclassified = [];

    for (let fi = 0; fi < fileDataArray.length; fi++) {
        const { name, data } = fileDataArray[fi];
        const pdfjs = await pdfjsLib.getDocument({ data: data.slice(0) }).promise;
        const numPages = pdfjs.numPages;

        // Use the dedicated cropping platform detection
        const detected = await detectCroppingPlatform(data);
        console.log(`[Cropping Auto] ${name} -> platform: ${detected.platform}, seller: ${detected.sellerId}`);

        if (detected.platform === 'meesho') {
            buckets.meesho.push({ name, data, numPages, sellerId: detected.sellerId });
        } else if (detected.platform === 'flipkart') {
            buckets.flipkart.push({ name, data, numPages, sellerId: detected.sellerId });
        } else if (detected.platform === 'amazon') {
            buckets.amazon.push({ name, data, numPages });
        } else {
            unclassified.push(name);
            console.warn(`[Cropping] Could not classify ${name}`);
        }

        const pct = 10 + Math.round((fi + 1) / fileDataArray.length * 20);
        updateCroppingProgress(pct, `Classified ${fi + 1}/${fileDataArray.length} files...`);
    }

    // Alert user about unclassified files
    if (unclassified.length > 0) {
        const msg = unclassified.length === fileDataArray.length
            ? `Could not identify the platform for any of the uploaded files. Please use the specific platform button (Meesho, Flipkart, or Amazon) instead.`
            : `Could not identify platform for: ${unclassified.join(', ')}. These files were skipped.`;
        alert(msg);
    }

    // If all files were unclassified, stop
    if (buckets.meesho.length === 0 && buckets.flipkart.length === 0 && buckets.amazon.length === 0) {
        updateCroppingProgress(100, '❌ No labels could be classified. Try using platform-specific buttons.');
        return;
    }

    // Process each bucket
    let progress = 30;
    const totalBuckets = (buckets.meesho.length > 0 ? 1 : 0) + (buckets.flipkart.length > 0 ? 1 : 0) + (buckets.amazon.length > 0 ? 1 : 0);
    const pctPerBucket = totalBuckets > 0 ? 60 / totalBuckets : 60;

    if (buckets.meesho.length > 0) {
        updateCroppingProgress(progress, 'Processing Meesho labels...');
        await processMeeshoBucket(buckets.meesho);
        progress += pctPerBucket;
    }
    if (buckets.flipkart.length > 0) {
        updateCroppingProgress(progress, 'Processing Flipkart labels...');
        await processFlipkartBucket(buckets.flipkart);
        progress += pctPerBucket;
    }
    if (buckets.amazon.length > 0) {
        updateCroppingProgress(progress, 'Processing Amazon labels...');
        await processAmazonBucket(buckets.amazon);
        progress += pctPerBucket;
    }

    updateCroppingProgress(100, '✅ Processing complete!');
}

// ------------------------------------------------------------------
// MANUAL MODE: treat all files as the selected platform, with validation
// ------------------------------------------------------------------
async function processCroppingManual(fileDataArray, platform) {
    updateCroppingProgress(10, `Validating labels for ${platform}...`);

    // First validate that all uploaded files actually belong to the selected platform
    const validFiles = [];
    const invalidFiles = [];

    for (const { name, data } of fileDataArray) {
        const detected = await detectCroppingPlatform(data);
        console.log(`[Cropping Manual] ${name} -> detected: ${detected.platform}, expected: ${platform}`);

        if (detected.platform === platform) {
            const pdfjs = await pdfjsLib.getDocument({ data: data.slice(0) }).promise;
            validFiles.push({ name, data, numPages: pdfjs.numPages, sellerId: detected.sellerId });
        } else {
            invalidFiles.push({ name, detectedPlatform: detected.platform });
        }
    }

    // Show error popup for invalid files
    if (invalidFiles.length > 0) {
        const platformName = platform.charAt(0).toUpperCase() + platform.slice(1);
        const fileDetails = invalidFiles.map(f => {
            const detName = f.detectedPlatform === 'unknown' ? 'Unknown' : f.detectedPlatform.charAt(0).toUpperCase() + f.detectedPlatform.slice(1);
            return `• ${f.name} (detected as ${detName})`;
        }).join('\n');

        if (validFiles.length === 0) {
            alert(`❌ Invalid Labels!\n\nNone of the uploaded files are ${platformName} labels.\n\n${fileDetails}\n\nPlease upload valid ${platformName} labels or use the Automatic button.`);
            updateCroppingProgress(100, `❌ No valid ${platformName} labels found.`);
            return;
        } else {
            alert(`⚠️ Some files are not ${platformName} labels and will be skipped:\n\n${fileDetails}`);
        }
    }

    if (validFiles.length === 0) {
        updateCroppingProgress(100, `❌ No valid ${platform} labels to process.`);
        return;
    }

    updateCroppingProgress(20, `Processing ${validFiles.length} ${platform} label(s)...`);

    if (platform === 'meesho') {
        await processMeeshoBucket(validFiles);
    } else if (platform === 'flipkart') {
        await processFlipkartBucket(validFiles);
    } else if (platform === 'amazon') {
        await processAmazonBucket(validFiles);
    }

    updateCroppingProgress(100, '✅ Processing complete!');
}

// ==================================================================
// SELLER ID EXTRACTION HELPERS FOR CROPPING
// ==================================================================

/**
 * Extract Meesho seller ID (return address) from text content items.
 * Looks for "If undelivered, return to: SELLERNAME" pattern.
 */
function extractMeeshoSellerId(textContentItems) {
    const items = textContentItems;
    const fullText = items.map(i => i.str).join(' ');
    
    // Try multiple patterns
    const patterns = [
        /if\s+undelivered,?\s+return\s+to:?\s*([A-Za-z0-9_\-]+)/i,
        /return\s+to:?\s*([A-Za-z0-9_\-]+)/i,
        /return\s+address:?\s*([A-Za-z0-9_\-]+)/i
    ];
    
    for (const pattern of patterns) {
        const match = fullText.match(pattern);
        if (match && match[1]) {
            const sellerId = match[1].trim();
            console.log(`[Meesho ID] Extracted seller ID: "${sellerId}"`);
            return sellerId;
        }
    }
    
    // Positional approach: find "return to" text and read next items
    for (let i = 0; i < items.length; i++) {
        const s = items[i].str.toLowerCase();
        if (s.includes('return to') || s.includes('return address')) {
            // Look at next few items for the seller ID
            for (let j = i + 1; j < Math.min(i + 5, items.length); j++) {
                const candidate = items[j].str.trim();
                if (!candidate) continue;
                // Seller ID pattern: alphanumeric with hyphens/underscores, 3+ chars
                if (/^[A-Za-z][A-Za-z0-9_\-]{2,}$/i.test(candidate)) {
                    console.log(`[Meesho ID] Extracted seller ID (positional): "${candidate}"`);
                    return candidate;
                }
            }
        }
    }
    
    console.log('[Meesho ID] Could not extract seller ID');
    return 'UNKNOWN';
}

// ==================================================================
// SKU EXTRACTION FUNCTIONS
// ==================================================================

/**
 * Extract SKU and Qty from Meesho label text items.
 * Pattern: "SKU" label above the SKU value, "Qty" label above the quantity in a table.
 * Uses positional matching (X/Y coordinates) to find values below headers.
 */
function extractMeeshoSKU(textContentItems) {
    const items = textContentItems;
    const results = [];

    // Helper to extract X,Y from transform
    const getPosition = (item) => {
        if (item.transform && item.transform.length >= 6) {
            return { x: item.transform[4], y: item.transform[5] };
        }
        return { x: 0, y: 0 };
    };

    // Find SKU headers and their positions
    const skuHeaders = [];
    const qtyHeaders = [];

    for (let i = 0; i < items.length; i++) {
        const s = items[i].str.trim();
        if (/^SKU$/i.test(s)) {
            const pos = getPosition(items[i]);
            skuHeaders.push({ idx: i, pos });
            console.log(`[Meesho SKU] Found SKU header at index ${i}, position (${pos.x.toFixed(1)}, ${pos.y.toFixed(1)})`);
        }
        if (/^Qty$/i.test(s) || /^Quantity$/i.test(s)) {
            qtyHeaders.push({ idx: i, pos: getPosition(items[i]) });
        }
    }

    // For each SKU header, find ALL values directly below it (multiple rows)
    for (const skuHeader of skuHeaders) {
        const skuCandidates = []; // Array of { sku, yPos }

        // Find all text items in the same column (similar X) but below (lower Y)
        // In PDF coordinates, Y increases upward, so "below" means lower Y value
        for (let j = 0; j < items.length; j++) {
            if (j === skuHeader.idx) continue;
            const candidate = items[j].str.trim();
            if (!candidate || candidate.length < 3) continue; // Minimum 3 chars for SKU
            
            // Skip if it's another header or common label text
            if (/^(SKU|Size|Qty|Color|Order|Price|Amount|Tax|HSN|Description|Product|Details|Name|Address|Phone|Mobile|City|State|Pin|Pincode|ZIP|Customer|Delivery|Ship|To|From)$/i.test(candidate)) continue;

            const pos = getPosition(items[j]);
            const xDiff = Math.abs(pos.x - skuHeader.pos.x);
            const yDiff = skuHeader.pos.y - pos.y; // positive if below

            // Item is in same column (small X difference) and below (positive Y diff)
            // Reduced Y threshold from 300 to 120 to only capture product table, not address section
            if (xDiff < 50 && yDiff > 0 && yDiff < 120) {
                // SKU pattern: must contain BOTH letters AND numbers (mixed), with optional hyphens/underscores
                if (/^[a-z0-9][a-z0-9_\-]{2,}$/i.test(candidate) && /[a-z]/i.test(candidate) && /[0-9]/.test(candidate)) {
                    skuCandidates.push({ sku: candidate, yPos: pos.y });
                    console.log(`[Meesho SKU] Found SKU candidate: "${candidate}" at Y: ${pos.y.toFixed(1)}`);
                }
            }
        }

        // For each SKU candidate, find its quantity value
        // Quantity should be in the Qty column (different X) but same row (similar Y)
        for (const skuCandidate of skuCandidates) {
            let qtyValue = '1';
            
            for (const qtyHeader of qtyHeaders) {
                let bestQtyCandidate = null;
                let minYDiff = Infinity;

                for (let j = 0; j < items.length; j++) {
                    if (j === qtyHeader.idx) continue;
                    const candidate = items[j].str.trim();
                    if (!candidate || !/^\d+$/.test(candidate)) continue;

                    const pos = getPosition(items[j]);
                    const xDiff = Math.abs(pos.x - qtyHeader.pos.x);
                    const yDiff = Math.abs(pos.y - skuCandidate.yPos); // same row = similar Y

                    // Item is in Qty column (same X as header) and same row as SKU (similar Y)
                    if (xDiff < 50 && yDiff < 10) {
                        if (yDiff < minYDiff) {
                            minYDiff = yDiff;
                            bestQtyCandidate = candidate;
                        }
                    }
                }

                if (bestQtyCandidate) {
                    qtyValue = bestQtyCandidate;
                    break;
                }
            }

            results.push({ sku: skuCandidate.sku, qty: qtyValue });
        }
    }

    // Fallback: sequential search (original logic)
    if (results.length === 0) {
        for (let i = 0; i < items.length; i++) {
            const s = items[i].str.trim();
            if (/^SKU$/i.test(s)) {
                for (let j = i + 1; j < Math.min(i + 10, items.length); j++) {
                    const candidate = items[j].str.trim();
                    if (!candidate) continue;
                    if (/^(Qty|Quantity|HSN|Description|Size|Color|Price)$/i.test(candidate)) break;
                    // SKU must contain BOTH letters AND numbers
                    if (candidate.length >= 2 && /[a-z]/i.test(candidate) && /[0-9]/.test(candidate)) {
                        const qtyMatch = items.slice(i, Math.min(i + 15, items.length))
                            .map(it => it.str.trim())
                            .find(str => /^\d+$/.test(str));
                        results.push({
                            sku: candidate,
                            qty: qtyMatch || '1'
                        });
                        break;
                    }
                }
                break;
            }
        }
    }

    // Last fallback: regex on joined text — SKU must contain both letters and numbers
    if (results.length === 0) {
        const fullText = items.map(i => i.str).join('\n');
        const skuMatch = fullText.match(/SKU\s*\n\s*([a-z0-9_\-]+)/i);
        const qtyMatch = fullText.match(/Qty\s*\n\s*(\d+)/i);
        if (skuMatch && /[a-z]/i.test(skuMatch[1]) && /[0-9]/.test(skuMatch[1])) {
            results.push({
                sku: skuMatch[1].trim(),
                qty: qtyMatch ? qtyMatch[1] : '1'
            });
        }
    }

    console.log(`[Meesho SKU] Extraction complete. Found ${results.length} SKU(s):`, results);
    return results;
}

/**
 * Extract SKU and Qty from Flipkart label text items.
 * Header: "SKU ID | Description", Qty column: "QTY"
 */
function extractFlipkartSKU(textContentItems, pageHeight = 842) {
    const items = textContentItems;
    const results = [];
    const fullText = items.map(i => i.str).join('\n');

    // Define SKU extraction bounding box (user-provided coordinates from top)
    // Convert from top coordinates to PDF bottom coordinates
    const SKU_BOX = {
        left: 191,
        right: 402,
        top: pageHeight - 279,    // y-279 from top
        bottom: pageHeight - 331  // y-331 from top
    };
    
    console.log(`[Flipkart SKU] Page height: ${pageHeight}, SKU box (PDF coords): left=${SKU_BOX.left}, right=${SKU_BOX.right}, bottom=${SKU_BOX.bottom.toFixed(1)}, top=${SKU_BOX.top.toFixed(1)}`);

    // Helper to extract X,Y from transform
    const getPosition = (item) => {
        if (item.transform && item.transform.length >= 6) {
            return { x: item.transform[4], y: item.transform[5] };
        }
        return { x: 0, y: 0 };
    };
    
    // Helper to check if position is within SKU box
    const isInSkuBox = (pos) => {
        return pos.x >= SKU_BOX.left && pos.x <= SKU_BOX.right &&
               pos.y >= SKU_BOX.bottom && pos.y <= SKU_BOX.top;
    };

    // Collect all text items within the SKU bounding box and log them
    const boxItems = [];
    for (let i = 0; i < items.length; i++) {
        const s = items[i].str.trim();
        if (!s) continue;
        const pos = getPosition(items[i]);
        if (!isInSkuBox(pos)) continue;
        boxItems.push({ text: s, x: pos.x, y: pos.y, idx: i });
        console.log(`[Flipkart SKU] Box item: "${s}" at (${pos.x.toFixed(1)}, ${pos.y.toFixed(1)})`);
    }

    // Group items into rows by Y position (items within 8pt are same row)
    const rows = [];
    const usedIndices = new Set();
    for (const item of boxItems) {
        if (usedIndices.has(item.idx)) continue;
        const row = [item];
        usedIndices.add(item.idx);
        for (const other of boxItems) {
            if (usedIndices.has(other.idx)) continue;
            if (Math.abs(other.y - item.y) < 8) {
                row.push(other);
                usedIndices.add(other.idx);
            }
        }
        // Sort items in row by X position (left to right)
        row.sort((a, b) => a.x - b.x);
        rows.push({ items: row, y: row[0].y, text: row.map(r => r.text).join(' ') });
    }

    // Sort rows by Y descending (top rows first in PDF coordinates)
    rows.sort((a, b) => b.y - a.y);
    console.log(`[Flipkart SKU] Found ${rows.length} rows in SKU box:`);
    rows.forEach((r, i) => console.log(`  Row ${i}: "${r.text}" at Y=${r.y.toFixed(1)}`));

    // Find header row (contains "SKU" keyword)
    let headerRowIdx = -1;
    for (let i = 0; i < rows.length; i++) {
        if (/SKU/i.test(rows[i].text)) {
            headerRowIdx = i;
            console.log(`[Flipkart SKU] Header row found at index ${i}: "${rows[i].text}"`);
            break;
        }
    }

    // Process data rows below the header
    if (headerRowIdx >= 0) {
        for (let i = headerRowIdx + 1; i < rows.length; i++) {
            const rowText = rows[i].text;
            // Skip empty or header-like rows
            if (/^(Total|Shipping|Amount|Invoice|Tax|IGST|CGST|SGST)/i.test(rowText)) continue;

            let skuValue = null;
            let qtyValue = '1';

            // Method 1: Row text contains "|" — take text before it
            if (rowText.includes('|')) {
                const beforePipe = rowText.split('|')[0].trim();
                // Remove leading row number like "1 " or "1."
                const cleaned = beforePipe.replace(/^\d+[\s.]+/, '').trim();
                if (cleaned.length >= 2) {
                    skuValue = cleaned;
                }
            }

            // Method 2: Look for "|" in individual items
            if (!skuValue) {
                for (const item of rows[i].items) {
                    if (item.text.includes('|')) {
                        const before = item.text.split('|')[0].trim();
                        const cleaned = before.replace(/^\d+[\s.]+/, '').trim();
                        if (cleaned.length >= 2) {
                            skuValue = cleaned;
                            break;
                        }
                    }
                }
            }

            // Method 3: If no "|" found, take the first alphanumeric token that looks like a SKU
            if (!skuValue) {
                for (const item of rows[i].items) {
                    const t = item.text.replace(/^\d+[\s.]+/, '').trim();
                    if (t.length >= 3 && /^[A-Za-z0-9][A-Za-z0-9_\-]{1,}$/.test(t)) {
                        skuValue = t;
                        break;
                    }
                }
            }

            // Find QTY: look for a standalone number in the row items (rightmost numeric item)
            for (let k = rows[i].items.length - 1; k >= 0; k--) {
                if (/^\d+$/.test(rows[i].items[k].text)) {
                    qtyValue = rows[i].items[k].text;
                    break;
                }
            }

            if (skuValue) {
                results.push({ sku: skuValue, qty: qtyValue });
                console.log(`[Flipkart SKU] Extracted SKU: "${skuValue}", Qty: ${qtyValue}`);
            }
        }
    }

    // Fallback: search all box text for SKU pattern
    if (results.length === 0) {
        const allBoxText = boxItems.map(b => b.text).join(' ');
        console.log(`[Flipkart SKU] Fallback: searching all box text: "${allBoxText}"`);
        
        // Try to find "text | description" pattern
        const pipeMatch = allBoxText.match(/(?:^|\s)([A-Za-z0-9][A-Za-z0-9_\-]+)\s*\|/);
        if (pipeMatch) {
            const sku = pipeMatch[1].trim();
            // Find qty nearby
            const qtyMatch = allBoxText.match(/(\d+)\s*$/);
            results.push({ sku: sku, qty: qtyMatch ? qtyMatch[1] : '1' });
            console.log(`[Flipkart SKU] Fallback found SKU: "${sku}"`);
        }
    }
    
    // Last resort fallback: regex on entire page joined text
    if (results.length === 0) {
        const skuMatch = fullText.match(/SKU\s*ID\s*\|?\s*Description[^\n]*\n[^\n]*?([A-Za-z0-9][A-Za-z0-9_\-]+)\s*\|/i);
        const qtyMatch = fullText.match(/QTY\s*\n\s*(\d+)/i);
        if (skuMatch) {
            results.push({
                sku: skuMatch[1].trim(),
                qty: qtyMatch ? qtyMatch[1] : '1'
            });
        }
    }

    console.log(`[Flipkart SKU] Extraction complete. Found ${results.length} SKU(s):`, results);
    return results;
}

/**
 * Extract SKU and Qty from Amazon invoice page text items.
 * Pattern: "( sku_value )" followed by "HSN:" on next line.
 * Qty is found in the table under "Qty" header.
 */
function extractAmazonSKU(textContentItems) {
    const items = textContentItems;
    const results = [];
    const fullText = items.map(i => i.str).join('\n');

    console.log('[Amazon SKU] Starting extraction...');

    // Helper to extract position
    const getPosition = (item) => {
        if (item.transform && item.transform.length >= 6) {
            return { x: item.transform[4], y: item.transform[5] };
        }
        return { x: 0, y: 0 };
    };

    // Group text items into rows by Y position (within 5pt = same row)
    const rowMap = new Map();
    for (let i = 0; i < items.length; i++) {
        const s = items[i].str.trim();
        if (!s) continue;
        const pos = getPosition(items[i]);
        const yKey = Math.round(pos.y / 5) * 5; // bucket by 5pt
        if (!rowMap.has(yKey)) rowMap.set(yKey, []);
        rowMap.set(yKey, [...rowMap.get(yKey), { text: s, x: pos.x, y: pos.y, idx: i }]);
    }

    // Build rows sorted by Y descending (top first in PDF coords)
    const rows = [];
    for (const [yKey, rowItems] of rowMap.entries()) {
        rowItems.sort((a, b) => a.x - b.x);
        const rowText = rowItems.map(r => r.text).join(' ');
        rows.push({ items: rowItems, y: yKey, text: rowText });
    }
    rows.sort((a, b) => b.y - a.y);

    // Method 1: Find rows containing "( sku )" followed by "HSN" — SKU must have NO spaces
    for (const row of rows) {
        const rowText = row.text;
        // Match all bracket patterns where content has NO spaces
        const bracketMatches = [...rowText.matchAll(/\(\s*(\S+)\s*\)/g)];
        for (const match of bracketMatches) {
            const candidate = match[1].trim();
            // Verify HSN appears in same row or nearby text
            if (/HSN/i.test(rowText) && candidate.length >= 2) {
                const alreadyFound = results.some(r => r.sku === candidate);
                if (!alreadyFound) {
                    results.push({ sku: candidate, qty: '1' });
                    console.log(`[Amazon SKU] Found SKU: "${candidate}" in row: "${rowText.substring(0, 80)}..."`);
                }
            }
        }
    }

    // Method 2: If HSN is on next row, use previous row bracket content as SKU
    for (let i = 0; i < rows.length - 1; i++) {
        const rowText = rows[i].text;
        const nextRowText = rows[i + 1].text;
        if (!/HSN/i.test(nextRowText)) continue;

        const bracketMatches = [...rowText.matchAll(/\(\s*(\S+)\s*\)/g)];
        for (const match of bracketMatches) {
            const candidate = match[1].trim();
            if (candidate.length < 2) continue;
            const alreadyFound = results.some(r => r.sku === candidate);
            if (!alreadyFound) {
                results.push({ sku: candidate, qty: '1' });
                console.log(`[Amazon SKU] Found SKU from adjacent HSN row: "${candidate}"`);
            }
        }
    }

    // Extract quantities with layered strategy
    const qtyValues = [];
    const addQtyValue = (val, y = Number.NEGATIVE_INFINITY) => {
        const normalized = String(val || '').trim();
        if (!/^\d{1,3}$/.test(normalized)) return;
        const numericVal = parseInt(normalized, 10);
        if (numericVal <= 0 || numericVal > 500) return;
        const exists = qtyValues.some(q => q.val === normalized && Math.abs((q.y || 0) - (y || 0)) < 2);
        if (!exists) qtyValues.push({ val: normalized, y });
    };

    // Strategy A: parse invoice table rows under header row containing Description + Qty
    const headerRow = rows.find(r => /Description/i.test(r.text) && /\bQty\b/i.test(r.text));
    if (headerRow) {
        const qtyHeaderItem = headerRow.items.find(it => /^Qty$/i.test(it.text) || /^Quantity$/i.test(it.text));
        if (qtyHeaderItem) {
            const qtyX = qtyHeaderItem.x;
            for (const row of rows) {
                if (row.y >= headerRow.y - 3) continue; // only rows below header
                if (/^(TOTAL|Amount\s+in\s+Words|Tax\s+Amount|For\s+)/i.test(row.text)) break;

                const numericItems = row.items.filter(it => /^\d{1,3}$/.test(it.text));
                if (numericItems.length === 0) continue;

                let best = null;
                for (const numItem of numericItems) {
                    const xDiff = Math.abs(numItem.x - qtyX);
                    if (xDiff <= 45 && (!best || xDiff < best.xDiff)) {
                        best = { value: numItem.text, xDiff };
                    }
                }

                if (best) addQtyValue(best.value, row.y);
            }
        }
    }

    // Strategy B: positional fallback from Qty header token
    const qtyHeaderItem = items.find(i => /^Qty$/i.test((i.str || '').trim()) || /^Quantity$/i.test((i.str || '').trim()));
    if (qtyHeaderItem) {
        const qtyHeaderPos = getPosition(qtyHeaderItem);
        for (let i = 0; i < items.length; i++) {
            const s = (items[i].str || '').trim();
            if (!/^\d{1,3}$/.test(s)) continue;
            const pos = getPosition(items[i]);
            const xDiff = Math.abs(pos.x - qtyHeaderPos.x);
            const yDiff = qtyHeaderPos.y - pos.y;
            if (xDiff < 45 && yDiff > 5 && yDiff < 450) {
                addQtyValue(s, pos.y);
            }
        }
    }

    // Strategy C: text fallback near Qty label
    const qtyAfterHeader = fullText.match(/Qty\s*[\n\r:\-]*\s*(\d{1,3})\b/i);
    if (qtyAfterHeader) {
        addQtyValue(qtyAfterHeader[1]);
    }

    // Strategy D: table-line fallback near HSN value (common Amazon invoice format)
    const qtyNearHsn = fullText.match(/HSN\s*:?\s*\d{4,}[\s\S]{0,80}?\b(\d{1,3})\b[\s\S]{0,30}?₹/i);
    if (qtyNearHsn) {
        addQtyValue(qtyNearHsn[1]);
    }

    qtyValues.sort((a, b) => b.y - a.y);

    // Assign quantities to results (match by order)
    for (let k = 0; k < results.length; k++) {
        if (k < qtyValues.length) {
            results[k].qty = qtyValues[k].val;
        }
    }

    // Fallback: regex on full text — bracket content with NO spaces, followed by HSN
    if (results.length === 0) {
        const allMatches = [...fullText.matchAll(/\(\s*(\S+)\s*\)\s*[\n\s]*HSN/gi)];
        for (const m of allMatches) {
            const candidate = m[1].trim();
            if (candidate.length >= 2) {
                results.push({ sku: candidate, qty: '1' });
                console.log(`[Amazon SKU] Fallback found SKU: "${candidate}"`);
            }
        }
        // Try to get Qty
        const qtyFallback = fullText.match(/Qty\s*\n?\s*(\d{1,3})/i);
        if (qtyFallback && results.length > 0) {
            results[0].qty = qtyFallback[1];
        }
    }

    // If SKU was not extractable but quantity exists, keep quantity for multi-qty detection
    if (results.length === 0 && qtyValues.length > 0) {
        results.push({ sku: '', qty: qtyValues[0].val });
        console.log(`[Amazon SKU] Quantity-only fallback: Qty ${qtyValues[0].val}`);
    }

    console.log(`[Amazon SKU] Extraction complete. Found ${results.length} SKU(s):`, results);
    return results;
}

/**
 * Extract order quantity directly from Amazon invoice page text.
 * Returns a numeric quantity or null if not confidently detected.
 */
function extractAmazonInvoiceQuantity(textContentItems) {
    if (!Array.isArray(textContentItems) || textContentItems.length === 0) {
        return null;
    }

    const items = textContentItems
        .map(item => {
            const text = (item && item.str ? String(item.str).trim() : '');
            const transform = item && item.transform ? item.transform : null;
            return {
                text,
                x: transform && transform.length >= 6 ? transform[4] : 0,
                y: transform && transform.length >= 6 ? transform[5] : 0,
            };
        })
        .filter(item => item.text.length > 0);

    const fullText = items.map(item => item.text).join('\n');
    const compactText = fullText.replace(/\s+/g, ' ');
    const candidates = [];

    const addCandidate = (value, weight = 1) => {
        const qty = parseInt(String(value || '').trim(), 10);
        if (!Number.isFinite(qty) || qty < 1 || qty > 99) return;
        candidates.push({ qty, weight });
    };

    const isQtyHeaderText = (text) => {
        const normalized = String(text || '').replace(/[^a-z0-9]/gi, '').toLowerCase();
        return normalized === 'qty' || normalized === 'quantity';
    };

    // Strategy 1: Qty column positional extraction from header token
    const qtyHeaderItem = items.find(item => isQtyHeaderText(item.text));
    if (qtyHeaderItem) {
        for (const item of items) {
            if (!/^\d{1,2}$/.test(item.text)) continue;
            const xDiff = Math.abs(item.x - qtyHeaderItem.x);
            const yDiff = qtyHeaderItem.y - item.y;
            if (xDiff <= 70 && yDiff > 5 && yDiff < 500) {
                addCandidate(item.text, 4);
            }
        }
    }

    // Strategy 2: Table row extraction (Description + Qty header row, then rows below)
    const rowMap = new Map();
    for (const item of items) {
        const yKey = Math.round(item.y / 4) * 4;
        if (!rowMap.has(yKey)) rowMap.set(yKey, []);
        rowMap.get(yKey).push(item);
    }

    const rows = Array.from(rowMap.entries()).map(([yKey, rowItems]) => {
        const sorted = [...rowItems].sort((a, b) => a.x - b.x);
        return {
            y: yKey,
            items: sorted,
            text: sorted.map(i => i.text).join(' ')
        };
    }).sort((a, b) => b.y - a.y);

    const tableHeaderRow = rows.find(r => /description/i.test(r.text) && /(qty|quantity|q\s*ty)/i.test(r.text));
    if (tableHeaderRow) {
        const qtyHeaderInRow = tableHeaderRow.items.find(i => isQtyHeaderText(i.text));
        const qtyX = qtyHeaderInRow ? qtyHeaderInRow.x : null;

        for (const row of rows) {
            if (row.y >= tableHeaderRow.y - 2) continue;
            if (/^(TOTAL|Amount\s+in\s+Words|Tax\s+Amount|For\s+)/i.test(row.text)) break;

            const integerTokens = row.items.filter(i => /^\d{1,2}$/.test(i.text));
            if (integerTokens.length === 0) continue;

            const hasAmountLikeToken = row.items.some(i => /₹|^\d+[\d,]*\.\d+$/.test(i.text));
            if (!hasAmountLikeToken) continue;

            if (qtyX !== null) {
                let best = null;
                for (const token of integerTokens) {
                    const xDiff = Math.abs(token.x - qtyX);
                    if (xDiff <= 70 && (!best || xDiff < best.xDiff)) {
                        best = { qty: token.text, xDiff };
                    }
                }
                if (best) addCandidate(best.qty, 5);
            } else if (integerTokens.length === 1) {
                addCandidate(integerTokens[0].text, 3);
            }
        }
    }

    // Strategy 3: Regex fallback where Qty is between Unit Price and Net Amount values
    const betweenAmountsRegex = /(?:₹|rs\.?|inr)\s*\d[\d,]*(?:\.\d+)?\s+(\d{1,2})\s+(?:₹|rs\.?|inr)\s*\d[\d,]*(?:\.\d+)?/gi;
    for (const match of compactText.matchAll(betweenAmountsRegex)) {
        addCandidate(match[1], 6);
    }

    // Strategy 4: Direct Qty label regex fallback
    const qtyRegex = /\bq\s*ty\b\s*[:\-]?\s*(\d{1,2})\b/gi;
    for (const match of compactText.matchAll(qtyRegex)) {
        addCandidate(match[1], 2);
    }

    if (candidates.length === 0) {
        return null;
    }

    // Choose strongest candidate, preferring higher confidence and then larger qty when tied
    candidates.sort((a, b) => (b.weight - a.weight) || (b.qty - a.qty));
    const detectedQty = candidates[0].qty;
    console.log(`[Amazon Qty] Detected quantity: ${detectedQty}`, candidates);
    return detectedQty;
}

// Format SKU data as overlay string
function formatSKUOverlay(skuData) {
    if (!skuData || skuData.length === 0) return '';
    return skuData.map(s => `SKU: ${s.sku} | Qty: ${s.qty}`).join('\n');
}

// ==================================================================
// PDF LAYOUT TRANSFORMATION FUNCTIONS
// ==================================================================

/**
 * MEESHO: No layout change. Scale to 4x6, overlay SKU at bottom.
 */
async function processMeeshoBucket(bucket) {
    // Group by sellerId
    const groups = {};
    for (const item of bucket) {
        const key = item.sellerId || 'UNKNOWN';
        if (!groups[key]) groups[key] = [];
        groups[key].push(item);
    }

    for (const [sellerId, items] of Object.entries(groups)) {
        const outputPdf = await PDFLib.PDFDocument.create();
        const helveticaBold = await outputPdf.embedFont(PDFLib.StandardFonts.HelveticaBold);
        let labelNumber = 0;
        let totalPages = 0;

        for (const item of items) {
            const srcPdf = await PDFLib.PDFDocument.load(item.data.slice(0));
            const srcPdfjs = await pdfjsLib.getDocument({ data: item.data.slice(0) }).promise;

            for (let p = 0; p < srcPdf.getPageCount(); p++) {
                const srcPage = srcPdf.getPage(p);
                const { width: srcW, height: srcH } = srcPage.getSize();

                const page = await srcPdfjs.getPage(p + 1);
                const tc = await page.getTextContent();
                const textItems = (tc.items || []).filter(i => i && typeof i.str === 'string' && i.str.trim().length > 0);
                const joinedText = textItems.map(i => i.str).join(' ');

                // Case-sensitive anchor: process only pages containing exact "TAX INVOICE"
                if (!joinedText.includes('TAX INVOICE')) {
                    console.log(`[Meesho] Skipping page ${p + 1} of ${item.name} - TAX INVOICE not found (case-sensitive)`);
                    continue;
                }

                // Find TAX INVOICE y-range in PDF coordinates
                let taxTop = null;
                let taxBottom = null;

                // Direct item match first
                for (const itemText of textItems) {
                    const raw = String(itemText.str || '');
                    if (raw.includes('TAX INVOICE')) {
                        const y = itemText.transform ? itemText.transform[5] : 0;
                        const h = itemText.height || (itemText.transform ? Math.abs(itemText.transform[3]) : 10) || 10;
                        taxBottom = y;
                        taxTop = y + h;
                        break;
                    }
                }

                // Row fallback if TAX and INVOICE are separate tokens
                if (taxTop === null || taxBottom === null) {
                    const rows = [];
                    const rowTolerance = 2.5;
                    for (const t of textItems) {
                        const y = t.transform ? t.transform[5] : 0;
                        let row = rows.find(r => Math.abs(r.y - y) <= rowTolerance);
                        if (!row) {
                            row = { y, items: [] };
                            rows.push(row);
                        }
                        row.items.push(t);
                    }

                    for (const row of rows) {
                        row.items.sort((a, b) => {
                            const ax = a.transform ? a.transform[4] : 0;
                            const bx = b.transform ? b.transform[4] : 0;
                            return ax - bx;
                        });
                        const rowText = row.items.map(i => i.str).join(' ');
                        if (rowText.includes('TAX INVOICE')) {
                            let rowTop = -Infinity;
                            let rowBottom = Infinity;
                            for (const rt of row.items) {
                                const y = rt.transform ? rt.transform[5] : 0;
                                const h = rt.height || (rt.transform ? Math.abs(rt.transform[3]) : 10) || 10;
                                rowTop = Math.max(rowTop, y + h);
                                rowBottom = Math.min(rowBottom, y);
                            }
                            if (Number.isFinite(rowTop) && Number.isFinite(rowBottom)) {
                                taxTop = rowTop;
                                taxBottom = rowBottom;
                                break;
                            }
                        }
                    }
                }

                if (taxTop === null || taxBottom === null) {
                    console.log(`[Meesho] Skipping page ${p + 1} of ${item.name} - TAX INVOICE anchor position not resolved`);
                    continue;
                }

                // Determine invoice bottom by removing trailing white space to page bottom
                let minContentBottom = srcH;
                for (const t of textItems) {
                    const y = t.transform ? t.transform[5] : 0;
                    if (y < minContentBottom) minContentBottom = y;
                }

                // Crop bounds with small safety margins
                const margin = 6;
                const labelLeft = 0;
                const labelRight = srcW;
                const labelTop = srcH;
                const labelBottom = Math.max(0, Math.min(srcH - 1, taxBottom - margin));

                const invoiceLeft = 0;
                const invoiceRight = srcW;
                const invoiceTop = Math.min(srcH, Math.max(1, taxTop + margin));
                const invoiceBottom = Math.max(0, Math.min(invoiceTop - 1, minContentBottom - margin));

                const labelHeight = labelTop - labelBottom;
                const invoiceHeight = invoiceTop - invoiceBottom;

                if (labelHeight < 20 || invoiceHeight < 20) {
                    console.log(`[Meesho] Skipping page ${p + 1} of ${item.name} - invalid crop heights (label=${labelHeight.toFixed(1)}, invoice=${invoiceHeight.toFixed(1)})`);
                    continue;
                }

                labelNumber++;
                const numText = `#${labelNumber}`;
                const numFontSize = 14;

                const embeddedLabel = await outputPdf.embedPage(srcPage, {
                    left: labelLeft,
                    bottom: labelBottom,
                    right: labelRight,
                    top: labelTop,
                });

                const embeddedInvoice = await outputPdf.embedPage(srcPage, {
                    left: invoiceLeft,
                    bottom: invoiceBottom,
                    right: invoiceRight,
                    top: invoiceTop,
                });

                // PAGE 1: LABEL (rotated 90°, full page fit)
                const labelPage = outputPdf.addPage([288, 432]);
                const labelWidth = labelRight - labelLeft;
                const labelScaleW = 288 / labelHeight;
                const labelScaleH = 432 / labelWidth;
                const labelScale = Math.min(labelScaleW, labelScaleH);
                const rotatedLabelVisualW = labelHeight * labelScale;
                const rotatedLabelVisualH = labelWidth * labelScale;
                const labelX = (288 - rotatedLabelVisualW) / 2 + rotatedLabelVisualW;
                const labelY = (432 - rotatedLabelVisualH) / 2;

                labelPage.drawPage(embeddedLabel, {
                    x: labelX,
                    y: labelY,
                    width: labelWidth * labelScale,
                    height: labelHeight * labelScale,
                    rotate: PDFLib.degrees(90),
                });

                const labelNumWidth = helveticaBold.widthOfTextAtSize(numText, numFontSize);
                const labelNumX = 288 - labelNumWidth - 5;
                const labelNumY = 432 - numFontSize - 3;
                const numPadX = 3;
                const numPadY = 2;
                labelPage.drawRectangle({
                    x: Math.max(0, labelNumX - numPadX),
                    y: Math.max(0, labelNumY - numPadY),
                    width: Math.min(288, labelNumWidth + numPadX * 2),
                    height: Math.min(432, numFontSize + numPadY * 2),
                    color: PDFLib.rgb(1, 1, 1),
                });
                labelPage.drawText(numText, {
                    x: labelNumX,
                    y: labelNumY,
                    size: numFontSize,
                    font: helveticaBold,
                    color: PDFLib.rgb(0, 0, 0),
                });

                // PAGE 2: INVOICE (rotated 90°, full page fit)
                const invoicePage = outputPdf.addPage([288, 432]);
                const invoiceWidth = invoiceRight - invoiceLeft;
                const invScaleW = 288 / invoiceHeight;
                const invScaleH = 432 / invoiceWidth;
                const invScale = Math.min(invScaleW, invScaleH);
                const rotatedInvVisualW = invoiceHeight * invScale;
                const rotatedInvVisualH = invoiceWidth * invScale;
                const invX = (288 - rotatedInvVisualW) / 2 + rotatedInvVisualW;
                const invY = (432 - rotatedInvVisualH) / 2;

                invoicePage.drawPage(embeddedInvoice, {
                    x: invX,
                    y: invY,
                    width: invoiceWidth * invScale,
                    height: invoiceHeight * invScale,
                    rotate: PDFLib.degrees(90),
                });

                const invNumWidth = helveticaBold.widthOfTextAtSize(numText, numFontSize);
                const invNumX = 288 - invNumWidth - 5;
                const invNumY = 432 - numFontSize - 3;
                invoicePage.drawRectangle({
                    x: Math.max(0, invNumX - numPadX),
                    y: Math.max(0, invNumY - numPadY),
                    width: Math.min(288, invNumWidth + numPadX * 2),
                    height: Math.min(432, numFontSize + numPadY * 2),
                    color: PDFLib.rgb(1, 1, 1),
                });
                invoicePage.drawText(numText, {
                    x: invNumX,
                    y: invNumY,
                    size: numFontSize,
                    font: helveticaBold,
                    color: PDFLib.rgb(0, 0, 0),
                });

                totalPages += 2;
            }
        }

        const pdfBytes = await outputPdf.save();
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        const filename = `Meesho_${sellerId}_${labelNumber}.pdf`;
        croppingResults.push({ filename, platform: 'meesho', blob, pageCount: labelNumber, sellerId });
    }
}

/**
 * FLIPKART: Separate label and invoice into different pages.
 * Label page (top portion above dashed line) is tightly cropped without white spaces.
 * Invoice page (bottom portion below dashed line) is scaled to fit.
 */
async function processFlipkartBucket(bucket) {
    const groups = {};
    for (const item of bucket) {
        const key = item.sellerId || 'UNKNOWN';
        if (!groups[key]) groups[key] = [];
        groups[key].push(item);
    }

    for (const [sellerId, items] of Object.entries(groups)) {
        const outputPdf = await PDFLib.PDFDocument.create();
        const helveticaBold = await outputPdf.embedFont(PDFLib.StandardFonts.HelveticaBold);
        let labelNumber = 0; // Sequential number for pairing label + invoice
        let totalPages = 0;

        for (const item of items) {
            const srcPdf = await PDFLib.PDFDocument.load(item.data.slice(0));
            const srcPdfjs = await pdfjsLib.getDocument({ data: item.data.slice(0) }).promise;

            for (let p = 0; p < srcPdf.getPageCount(); p++) {
                // Get page dimensions first
                const srcPage = srcPdf.getPage(p);
                const { width: srcW, height: srcH } = srcPage.getSize();
                
                // Extract text and SKU (with page height for coordinate conversion)
                const page = await srcPdfjs.getPage(p + 1);
                const tc = await page.getTextContent();
                const skuData = extractFlipkartSKU(tc.items, srcH);

                // Process only pages that look like actual label pages (must contain barcode-like content)
                const hasBarcodeLikeContent = hasFlipkartBarcodeLikeContent(tc.items);
                if (!hasBarcodeLikeContent) {
                    console.log(`[Flipkart] Skipping page ${p + 1} of ${item.name} - no barcode-like content`);
                    continue;
                }

                // Increment label number for this pair
                labelNumber++;

                // ============================================================
                // EXACT CROP COORDINATES
                // Label: x-187 to x-406, y-25 to y-383
                // Invoice: x-29 to x-566, y-382 to y-800
                // ============================================================
                
                const labelLeft = 187;
                const labelRight = 406;
                const labelBottom = srcH - 383;
                const labelTop = srcH - 25;
                
                const invoiceLeft = 29;
                const invoiceRight = 566;
                const invoiceBottom = srcH - 800;
                const invoiceTop = srcH - 382;
                
                const labelWidth = labelRight - labelLeft;   // 219 points
                const labelHeight = labelTop - labelBottom;   // 358 points
                const invoiceWidth = invoiceRight - invoiceLeft; // 537 points
                const invoiceHeight = invoiceTop - invoiceBottom; // 418 points

                // Embed cropped regions
                const embeddedLabel = await outputPdf.embedPage(srcPage, {
                    left: labelLeft, bottom: labelBottom, right: labelRight, top: labelTop,
                });
                const embeddedInvoice = await outputPdf.embedPage(srcPage, {
                    left: invoiceLeft, bottom: invoiceBottom, right: invoiceRight, top: invoiceTop,
                });

                // ============================================================
                // PAGE 1: LABEL (not rotated)
                // ============================================================
                const labelPage = outputPdf.addPage([288, 432]);

                const labelScaleW = 288 / labelWidth;
                const labelScaleH = 432 / labelHeight;
                const labelScale = Math.min(labelScaleW, labelScaleH);

                const labelDrawW = labelWidth * labelScale;
                const labelDrawH = labelHeight * labelScale;

                // Position: centered, using full page area
                const labelX = (288 - labelDrawW) / 2;
                const labelY = (432 - labelDrawH) / 2;

                labelPage.drawPage(embeddedLabel, {
                    x: labelX,
                    y: labelY,
                    width: labelDrawW,
                    height: labelDrawH,
                });

                // Draw number at top-right of label page
                const numText = `#${labelNumber}`;
                const numFontSize = 14;
                const numTextWidth = helveticaBold.widthOfTextAtSize(numText, numFontSize);
                labelPage.drawText(numText, {
                    x: 288 - numTextWidth - 5,
                    y: 432 - numFontSize - 3,
                    size: numFontSize,
                    font: helveticaBold,
                    color: PDFLib.rgb(0, 0, 0),
                });

                // ============================================================
                // PAGE 2: INVOICE (rotated 90°)
                // ============================================================
                const invoicePage = outputPdf.addPage([288, 432]);

                // After 90° rotation: visual width = invoiceHeight, visual height = invoiceWidth
                const invScaleW = 288 / invoiceHeight;
                const invScaleH = 432 / invoiceWidth;
                const invScale = Math.min(invScaleW, invScaleH);

                const rotatedInvVisualW = invoiceHeight * invScale;
                const rotatedInvVisualH = invoiceWidth * invScale;

                const invX = (288 - rotatedInvVisualW) / 2 + rotatedInvVisualW;
                const invY = (432 - rotatedInvVisualH) / 2;

                invoicePage.drawPage(embeddedInvoice, {
                    x: invX,
                    y: invY,
                    width: invoiceWidth * invScale,
                    height: invoiceHeight * invScale,
                    rotate: PDFLib.degrees(90),
                });

                // Draw same number at top-right of invoice page
                const invNumTextWidth = helveticaBold.widthOfTextAtSize(numText, numFontSize);
                invoicePage.drawText(numText, {
                    x: 288 - invNumTextWidth - 5,
                    y: 432 - numFontSize - 3,
                    size: numFontSize,
                    font: helveticaBold,
                    color: PDFLib.rgb(0, 0, 0),
                });

                totalPages += 2; // Label page + Invoice page
            }
        }

        const pdfBytes = await outputPdf.save();
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        const filename = `Flipkart_${sellerId}_${labelNumber}.pdf`;
        croppingResults.push({ filename, platform: 'flipkart', blob, pageCount: labelNumber, sellerId });
    }
}

function hasFlipkartBarcodeLikeContent(textItems) {
    if (!textItems || !textItems.length) return false;

    const normalizedItems = textItems
        .map(item => (item && item.str ? String(item.str).trim() : ''))
        .filter(Boolean);

    if (normalizedItems.length === 0) return false;

    const joined = normalizedItems.join(' ').toUpperCase();
    if (joined.includes('BARCODE')) return true;

    // Common barcode text signatures found in PDF extracted text
    const starWrappedPattern = /\*[A-Z0-9\-]{6,}\*/;
    const compactCodePattern = /^[A-Z0-9\-]{10,}$/;
    const longNumericPattern = /^\d{10,}$/;

    for (const token of normalizedItems) {
        const compact = token.replace(/\s+/g, '');
        if (starWrappedPattern.test(compact)) return true;
        if (compactCodePattern.test(compact) && /\d/.test(compact)) return true;
        if (longNumericPattern.test(compact)) return true;
    }

    return false;
}

/**
 * Scan PDF page canvas for 1D barcode using ZXing.
 * Returns barcode text if found, null otherwise.
 */
/**
 * Extract a region from canvas
 */
function extractCanvasRegion(sourceCanvas, x, y, width, height) {
    const region = document.createElement('canvas');
    region.width = width;
    region.height = height;
    const ctx = region.getContext('2d');
    ctx.drawImage(sourceCanvas, x, y, width, height, 0, 0, width, height);
    return region;
}

/**
 * Enhance image contrast for better OCR
 */
function enhanceImageForBarcode(canvas) {
    const enhanced = document.createElement('canvas');
    enhanced.width = canvas.width;
    enhanced.height = canvas.height;
    const ctx = enhanced.getContext('2d');
    
    // Draw original
    ctx.drawImage(canvas, 0, 0);
    
    // Get image data and enhance contrast
    const imageData = ctx.getImageData(0, 0, enhanced.width, enhanced.height);
    const data = imageData.data;
    
    // Apply contrast enhancement
    const contrast = 1.5;
    const factor = (259 * (contrast + 255)) / (255 * (259 - contrast));
    
    for (let i = 0; i < data.length; i += 4) {
        data[i] = factor * (data[i] - 128) + 128;     // Red
        data[i + 1] = factor * (data[i + 1] - 128) + 128; // Green
        data[i + 2] = factor * (data[i + 2] - 128) + 128; // Blue
        // Alpha stays the same
    }
    
    ctx.putImageData(imageData, 0, 0);
    return enhanced;
}

/**
 * OCR scan of PDF page to extract AWB text (for image-based PDFs)
 */
async function ocrScanForAWB(pdfjsPage) {
    try {
        console.log('[OCR Scanner] Starting OCR for AWB extraction...');
        
        // Render page at high resolution for better OCR
        const viewport = pdfjsPage.getViewport({ scale: 2.5 });
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        canvas.width = viewport.width;
        canvas.height = viewport.height;

        await pdfjsPage.render({
            canvasContext: context,
            viewport: viewport
        }).promise;

        console.log(`[OCR Scanner] Rendered canvas: ${canvas.width}x${canvas.height}`);

        // Apply preprocessing for better OCR
        const enhancedCanvas = enhanceImageForBarcode(canvas);
        
        // Perform OCR on top 25% of page (where AWB typically is)
        const topRegion = extractCanvasRegion(enhancedCanvas, 0, 0, enhancedCanvas.width, enhancedCanvas.height / 4);
        
        const { data: { text } } = await Tesseract.recognize(
            topRegion.toDataURL('image/png'),
            'eng',
            {
                logger: m => {
                    if (m.status === 'recognizing text') {
                        console.log(`[OCR Scanner] Progress: ${Math.round(m.progress * 100)}%`);
                    }
                }
            }
        );

        console.log('[OCR Scanner] OCR complete, searching for AWB...');
        console.log('[OCR Scanner] Extracted text:', text.substring(0, 200));

        // Search for AWB pattern in OCR text
        const awbMatch = text.match(/AWB\s*[:\s]*([0-9]{6,})/i);
        if (awbMatch) {
            console.log(`[OCR Scanner] Found AWB: ${awbMatch[1]}`);
            return awbMatch[1];
        }

        // Try alternate patterns
        const altMatch = text.match(/\b([0-9]{10,15})\b/);
        if (altMatch) {
            console.log(`[OCR Scanner] Found potential AWB number: ${altMatch[1]}`);
            return altMatch[1];
        }

        console.log('[OCR Scanner] No AWB found in OCR text');
    } catch (error) {
        console.log('[OCR Scanner] OCR failed:', error.message);
    }
    
    return null;
}

/**
 * Extract AWB number from Amazon label page using OCR.
 */
async function extractAmazonAWB(pdfjsPage) {
    if (!pdfjsPage) {
        console.log('[Amazon AWB] No page provided');
        return null;
    }
    
    console.log('[Amazon AWB] Starting OCR scan for AWB...');
    const ocrAWB = await ocrScanForAWB(pdfjsPage);
    
    if (ocrAWB) {
        console.log(`[Amazon AWB] Found AWB from OCR: ${ocrAWB}`);
        return ocrAWB;
    }
    
    console.log('[Amazon AWB] No AWB found');
    return null;
}

/**
 * AMAZON: Portrait 4x6 layout with label (top) and invoice (bottom), both rotated 90°.
 * SKU/AWB overlaid on invoice area.
 */
async function processAmazonBucket(bucket) {
    const outputPdf = await PDFLib.PDFDocument.create();
    const helveticaBold = await outputPdf.embedFont(PDFLib.StandardFonts.HelveticaBold);
    let labelCount = 0;

    for (const item of bucket) {
        const srcPdf = await PDFLib.PDFDocument.load(item.data.slice(0));
        const srcPdfjs = await pdfjsLib.getDocument({ data: item.data.slice(0) }).promise;
        const numPages = srcPdf.getPageCount();

        // Process in pairs: page 1+2, 3+4, etc.
        for (let p = 0; p < numPages; p += 2) {
            const labelPageIdx = p;
            const invoicePageIdx = p + 1;

            let skuData = [];
            let invoiceDetectedQty = null;
            let awbNumber = null;

            // Extract AWB from label page (odd page) using OCR
            const labelPage = await srcPdfjs.getPage(labelPageIdx + 1);
            awbNumber = await extractAmazonAWB(labelPage);

            // Extract SKU from invoice page (even page)
            if (invoicePageIdx < numPages) {
                const invPage = await srcPdfjs.getPage(invoicePageIdx + 1);
                const tc = await invPage.getTextContent();
                skuData = extractAmazonSKU(tc.items);
                invoiceDetectedQty = extractAmazonInvoiceQuantity(tc.items);

                // Ensure quantity is available even if SKU extraction is partial/empty
                if (invoiceDetectedQty && skuData.length === 0) {
                    skuData = [{ sku: '', qty: String(invoiceDetectedQty) }];
                } else if (invoiceDetectedQty && skuData.length > 0) {
                    const firstQty = parseInt(skuData[0].qty, 10);
                    if (!Number.isFinite(firstQty) || firstQty < invoiceDetectedQty) {
                        skuData[0].qty = String(invoiceDetectedQty);
                    }
                }
            }

            // Get label page dimensions
            const labelSrcPage = srcPdf.getPage(labelPageIdx);
            const { width: labelW, height: labelH } = labelSrcPage.getSize();
            const embeddedLabel = await outputPdf.embedPage(labelSrcPage);

            labelCount++;
            const numText = `#${labelCount}`;
            const numFontSize = 14;

            // ============================================================
            // LABEL OUTPUT PAGE: no crop, scaled to fit with reserved bottom info area
            // ============================================================
            const labelOutPage = outputPdf.addPage([288, 432]);
            const infoSpace = 28; // reserved bottom space for SKU + AWB text
            const labelAvailH = 432 - infoSpace;
            const labelScale = Math.min(288 / labelW, labelAvailH / labelH);
            const labelDrawW = labelW * labelScale;
            const labelDrawH = labelH * labelScale;

            labelOutPage.drawPage(embeddedLabel, {
                x: (288 - labelDrawW) / 2,
                y: infoSpace + (labelAvailH - labelDrawH) / 2,
                width: labelDrawW,
                height: labelDrawH,
            });

            // Number overlay on label page
            const labelNumWidth = helveticaBold.widthOfTextAtSize(numText, numFontSize);
            labelOutPage.drawText(numText, {
                x: 288 - labelNumWidth - 5,
                y: 432 - numFontSize - 3,
                size: numFontSize,
                font: helveticaBold,
                color: PDFLib.rgb(0, 0, 0),
            });

            // Reserved bottom text: SKU/AWB or Multi Quantity Order marker (no overlay)
            const hasMultiSku = (skuData || []).length >= 2;
            const hasMultiQty =
                (skuData || []).some(s => parseInt(s && s.qty, 10) > 1) ||
                (Number.isFinite(invoiceDetectedQty) && invoiceDetectedQty > 1);
            const isMultiQtyOrder = hasMultiSku || hasMultiQty;

            const uniqueSkus = Array.from(new Set((skuData || []).map(s => s && s.sku).filter(Boolean)));
            const skuText = isMultiQtyOrder
                ? 'Multi Quantity Order'
                : (uniqueSkus.length > 0 ? `SKU: ${uniqueSkus.join(', ')}` : '');
            const awbText = awbNumber ? `AWB: ${awbNumber}` : '';
            const infoText = [skuText, awbText].filter(Boolean).join(' | ');

            if (infoText) {
                labelOutPage.drawText(infoText, {
                    x: 6,
                    y: 8,
                    size: 10,
                    font: helveticaBold,
                    color: PDFLib.rgb(0, 0, 0),
                });
            }

            // ============================================================
            // INVOICE OUTPUT PAGE: no crop, full-page fit, same number overlay
            // ============================================================
            if (invoicePageIdx < numPages) {
                const invoiceSrcPage = srcPdf.getPage(invoicePageIdx);
                const { width: invW, height: invH } = invoiceSrcPage.getSize();
                const embeddedInvoice = await outputPdf.embedPage(invoiceSrcPage);
                const invoiceOutPage = outputPdf.addPage([288, 432]);
                const invScale = Math.min(288 / invW, 432 / invH);
                const invDrawW = invW * invScale;
                const invDrawH = invH * invScale;

                invoiceOutPage.drawPage(embeddedInvoice, {
                    x: (288 - invDrawW) / 2,
                    y: (432 - invDrawH) / 2,
                    width: invDrawW,
                    height: invDrawH,
                });

                const invNumWidth = helveticaBold.widthOfTextAtSize(numText, numFontSize);
                invoiceOutPage.drawText(numText, {
                    x: 288 - invNumWidth - 5,
                    y: 432 - numFontSize - 3,
                    size: numFontSize,
                    font: helveticaBold,
                    color: PDFLib.rgb(0, 0, 0),
                });
            }
        }
    }

    const pdfBytes = await outputPdf.save();
    const blob = new Blob([pdfBytes], { type: 'application/pdf' });
    const filename = `Amazon_${labelCount}_labels.pdf`;
    croppingResults.push({ filename, platform: 'amazon', blob, pageCount: labelCount, sellerId: 'Amazon' });
}

// ==================================================================
// DISPLAY RESULTS
// ==================================================================
function displayCroppingResults() {
    const el = getCroppingElements();

    let totalProcessedPages = 0;
    const platforms = new Set();
    el.outputList.innerHTML = '';

    for (const result of croppingResults) {
        totalProcessedPages += result.pageCount;
        platforms.add(result.platform);

        const card = document.createElement('div');
        card.className = `cropping-output-card platform-${result.platform}`;
        card.innerHTML = `
            <div class="output-info">
                <span class="output-filename">📄 ${result.filename}</span>
                <span class="output-details">${result.platform.charAt(0).toUpperCase() + result.platform.slice(1)} • ${result.pageCount} label(s) • Seller: ${result.sellerId}</span>
            </div>
            <button class="btn-download-single" data-idx="${croppingResults.indexOf(result)}">⬇️ Download</button>
        `;
        el.outputList.appendChild(card);
    }

    // Bind individual download buttons
    el.outputList.querySelectorAll('.btn-download-single').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const idx = parseInt(e.target.dataset.idx);
            downloadSingleCropped(idx);
        });
    });

    el.totalProcessed.textContent = totalProcessedPages;
    el.platformCount.textContent = platforms.size;
    el.resultsSection.style.display = 'block';
}

// ==================================================================
// DOWNLOAD FUNCTIONS
// ==================================================================
function downloadSingleCropped(index) {
    const result = croppingResults[index];
    if (!result) return;
    const url = URL.createObjectURL(result.blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = result.filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 5000);
}

async function downloadAllCropped() {
    if (croppingResults.length === 0) {
        alert('No processed PDFs to download.');
        return;
    }

    for (let i = 0; i < croppingResults.length; i++) {
        downloadSingleCropped(i);
        // Small delay between downloads to avoid browser blocking
        if (i < croppingResults.length - 1) {
            await new Promise(resolve => setTimeout(resolve, 500));
        }
    }
}

// Cropping tab initialization runs through startApplicationOnce() after device verification.