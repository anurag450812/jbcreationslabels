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
let stockAlreadyDeducted = false; // Prevent duplicate stock deductions on multiple download clicks

// Label Counter specific variables
let counterUploadedFiles = [];

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

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    initializeApp();
});

function initializeApp() {
    // Display priority list (with default labels first)
    displayPriorityList();
    
    // Setup event listeners
    setupEventListeners();
    
    // Restore active tab from localStorage
    const savedTab = localStorage.getItem('activeTab');
    if (savedTab && (savedTab === 'sorting' || savedTab === 'counter')) {
        switchTab(savedTab);
    }
    
    // Load priority labels from cloud (async, will update UI when loaded)
    loadPriorityLabelsFromCloud();
    
    // Load app configuration from cloud (async, will update settings when loaded)
    loadAppConfigFromCloud();
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
            const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
            
            for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                const page = await pdf.getPage(pageNum);
                const textContent = await page.getTextContent();
                
                // Extract text with rotation support (handles 90°, 180°, 270° rotated labels)
                const textVariants = extractTextWithRotationSupport(textContent);
                
                // Try matching label info against all rotation variants
                // Pass raw items for item-level fallback detection (handles rotated labels)
                const rawItems = textContent.items;
                const labelInfo = extractLabelInfoMulti(textVariants, file.name, pageNum, rawItems);
                if (labelInfo) {
                    allLabelsData.push(labelInfo);
                } else {
                    // Page didn't match any criteria - track it with reason
                    // Use combined text from all variants for skip reason analysis
                    const combinedText = textVariants.join(' ');
                    const reason = determineSkipReason(combinedText);
                    skippedPages.push({
                        fileName: file.name,
                        pageNum: pageNum,
                        fileIndex: fileIdx,
                        reason: reason
                    });
                }
                
                totalPages++;
                const progress = ((fileIdx * 100 / counterUploadedFiles.length) + 
                    (pageNum / pdf.numPages * 100 / counterUploadedFiles.length));
                counterProgressFill.style.width = `${progress}%`;
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
        const { priorityPages, otherPages, matchedLabels, labelCounts } = sortPages(allPages);
        
        // Store label counts globally for use during download
        labelOccurrences = labelCounts;
        
        // Reset stock deduction flag for this new processing batch
        stockAlreadyDeducted = false;
        
        updateProgress(80, 'Creating sorted PDF...');
        
        // Create new PDF with sorted pages
        processedPDF = await createSortedPDF([...priorityPages, ...otherPages], uploadedFiles);
        
        updateProgress(100, 'Complete!');
        
        // Display results with label counts
        displayResults(allPages.length, priorityPages.length, otherPages.length, matchedLabels, allPages, labelCounts);
        
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
    
    // Categorize pages
    for (const page of pages) {
        let matched = false;
        let matchedLabel = null;
        let priorityIndex = Infinity;
        
        // Check if page text contains any priority label
        for (const label of PRIORITY_LABELS) {
            const labelLower = label.toLowerCase();
            const textLower = page.text;
            
            // Find the label in the text
            const index = textLower.indexOf(labelLower);
            if (index !== -1) {
                // Check the character immediately after the label
                const charAfterLabel = textLower.charAt(index + labelLower.length);
                
                // Only match if there's NO comma or digit after the label
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
            priorityPages.push({ ...page, priorityIndex, matchedLabel });
            matchedLabels.add(matchedLabel);
            // Count the occurrence
            labelCounts[matchedLabel] = (labelCounts[matchedLabel] || 0) + 1;
        } else {
            otherPages.push(page);
        }
    }
    
    // Sort priority pages by their priority index
    priorityPages.sort((a, b) => a.priorityIndex - b.priorityIndex);
    
    // Filter to only labels that were found
    const foundLabelCounts = {};
    for (const label of matchedLabels) {
        foundLabelCounts[label] = labelCounts[label];
    }
    
    return {
        priorityPages,
        otherPages,
        matchedLabels: Array.from(matchedLabels),
        labelCounts: foundLabelCounts // Return the counts
    };
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
                <span class="platform-breakdown-label">pages processed</span>
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
                <span class="platform-breakdown-label">pages processed</span>
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
                <span class="platform-breakdown-label">pages processed</span>
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
                <span class="platform-breakdown-label">pages processed</span>
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
        }
    }
    
    passwordInput.focus();
}

function hidePasswordModal() {
    passwordModal.style.display = 'none';
    passwordInput.value = '';
    passwordError.style.display = 'none';
}

function checkPassword() {
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
        
        // After successful download, deduct stock from Google Sheets (only once per processing)
        if (Object.keys(labelOccurrences).length > 0 && GOOGLE_SHEETS_CONFIG.webAppUrl && !stockAlreadyDeducted) {
            const result = await deductStockFromGoogleSheets(labelOccurrences);
            if (result.success) {
                console.log('Stock deducted successfully');
                stockAlreadyDeducted = true; // Mark as deducted to prevent duplicate deductions
            } else {
                console.log('Stock deduction note:', result.message);
            }
        } else if (stockAlreadyDeducted) {
            console.log('Stock already deducted for this batch - skipping duplicate deduction');
            if (stockUpdateStatus) {
                stockUpdateStatus.innerHTML = '<span class="status-success">✅ Stock already updated for this batch</span>';
                stockUpdateStatus.style.display = 'block';
            }
        }
        
    } catch (error) {
        console.error('Error downloading PDF:', error);
        alert('Error downloading PDF. Please try again.');
    }
}

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

