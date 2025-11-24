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
    'sv24'
];

// Password for editing
const EDIT_PASSWORD = '200274';

// Global variables
let uploadedFiles = [];
let processedPDF = null;

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

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    initializeApp();
});

function initializeApp() {
    // Display priority list
    displayPriorityList();
    
    // Setup event listeners
    setupEventListeners();
}

function displayPriorityList() {
    priorityList.innerHTML = PRIORITY_LABELS.map(label => 
        `<div class="priority-item">${label}</div>`
    ).join('');
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
    editLabelsBtn.addEventListener('click', showPasswordModal);
    
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
}

function clearFiles() {
    uploadedFiles = [];
    fileInput.value = '';
    processBtn.disabled = true;
    clearBtn.style.display = 'none';
    platformInfo.style.display = 'none';
    resultsSection.style.display = 'none';
    statusSection.style.display = 'none';
    
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

function handleFiles(files) {
    // Merge with existing files instead of replacing
    const existingFileNames = new Set(uploadedFiles.map(f => f.name));
    const newFiles = files.filter(f => !existingFileNames.has(f.name));
    
    // Add only new files to avoid duplicates
    uploadedFiles = [...uploadedFiles, ...newFiles];
    
    if (uploadedFiles.length > 0) {
        processBtn.disabled = false;
        clearBtn.style.display = 'block';
        
        // Detect platforms from filenames
        const platformCounts = detectPlatforms(uploadedFiles);
        
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
        
        // Display platform info
        displayPlatformInfo(platformCounts);
    }
}

function detectPlatforms(files) {
    const counts = {
        flipkart: 0,
        amazon: 0,
        meesho: 0,
        unknown: 0
    };
    
    files.forEach(file => {
        const fileName = file.name.toLowerCase();
        if (fileName.includes('flipkart')) {
            counts.flipkart++;
        } else if (fileName.includes('amazon')) {
            counts.amazon++;
        } else if (fileName.includes('meesho')) {
            counts.meesho++;
        } else {
            counts.unknown++;
        }
    });
    
    return counts;
}

function displayPlatformInfo(counts) {
    const platforms = [
        { name: 'Flipkart', key: 'flipkart', icon: '🛒' },
        { name: 'Amazon', key: 'amazon', icon: '📦' },
        { name: 'Meesho', key: 'meesho', icon: '🛍️' }
    ];
    
    let html = '';
    platforms.forEach(platform => {
        if (counts[platform.key] > 0) {
            html += `
                <div class="platform-stat">
                    <div class="platform-stat-icon">${platform.icon}</div>
                    <span class="platform-stat-name">${platform.name}</span>
                    <span class="platform-stat-count">${counts[platform.key]} file(s)</span>
                </div>
            `;
        }
    });
    
    if (counts.unknown > 0) {
        html += `
            <div class="platform-stat">
                <div class="platform-stat-icon">❓</div>
                <span class="platform-stat-name">Unknown</span>
                <span class="platform-stat-count">${counts.unknown} file(s)</span>
            </div>
        `;
    }
    
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
        const { priorityPages, otherPages, matchedLabels } = sortPages(allPages);
        
        updateProgress(80, 'Creating sorted PDF...');
        
        // Create new PDF with sorted pages
        processedPDF = await createSortedPDF([...priorityPages, ...otherPages], uploadedFiles);
        
        updateProgress(100, 'Complete!');
        
        // Display results
        displayResults(allPages.length, priorityPages.length, otherPages.length, matchedLabels, allPages);
        
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
    
    // Detect platform from filename
    const platform = detectPlatformFromFilename(file.name);
    
    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        const text = textContent.items.map(item => item.str).join(' ');
        
        pages.push({
            fileName: file.name,
            pageNumber: i,
            text: text.toLowerCase(),
            platform: platform,
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

function sortPages(pages) {
    const priorityPages = [];
    const otherPages = [];
    const matchedLabels = new Set();
    const priorityMap = new Map();
    
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
        } else {
            otherPages.push(page);
        }
    }
    
    // Sort priority pages by their priority index
    priorityPages.sort((a, b) => a.priorityIndex - b.priorityIndex);
    
    return {
        priorityPages,
        otherPages,
        matchedLabels: Array.from(matchedLabels)
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

function displayResults(total, matched, unmatched, labels, allPages) {
    // Update stats
    document.getElementById('totalPages').textContent = total;
    document.getElementById('matchedPages').textContent = matched;
    document.getElementById('unmatchedPages').textContent = unmatched;
    
    // Calculate platform breakdown
    const platformCounts = {
        flipkart: 0,
        amazon: 0,
        meesho: 0,
        unknown: 0
    };
    
    allPages.forEach(page => {
        if (platformCounts.hasOwnProperty(page.platform)) {
            platformCounts[page.platform]++;
        }
    });
    
    // Display platform breakdown
    const platformBreakdownStats = document.getElementById('platformBreakdownStats');
    const platforms = [
        { name: 'Flipkart', key: 'flipkart', icon: '🛒' },
        { name: 'Amazon', key: 'amazon', icon: '📦' },
        { name: 'Meesho', key: 'meesho', icon: '🛍️' }
    ];
    
    let breakdownHtml = '';
    platforms.forEach(platform => {
        if (platformCounts[platform.key] > 0) {
            breakdownHtml += `
                <div class="platform-breakdown-item">
                    <div class="platform-breakdown-icon">${platform.icon}</div>
                    <span class="platform-breakdown-name">${platform.name}</span>
                    <span class="platform-breakdown-count">${platformCounts[platform.key]}</span>
                    <span class="platform-breakdown-label">pages processed</span>
                </div>
            `;
        }
    });
    
    if (platformCounts.unknown > 0) {
        breakdownHtml += `
            <div class="platform-breakdown-item">
                <div class="platform-breakdown-icon">❓</div>
                <span class="platform-breakdown-name">Unknown</span>
                <span class="platform-breakdown-count">${platformCounts.unknown}</span>
                <span class="platform-breakdown-label">pages processed</span>
            </div>
        `;
    }
    
    platformBreakdownStats.innerHTML = breakdownHtml;
    
    // Display matched labels
    const labelChips = document.getElementById('labelChips');
    labelChips.innerHTML = labels.map(label => 
        `<span class="label-chip">${label}</span>`
    ).join('');
    
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
function showPasswordModal() {
    passwordModal.style.display = 'flex';
    passwordInput.value = '';
    passwordError.style.display = 'none';
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
        showEditLabelsModal();
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

function saveLabels() {
    const newLabels = labelsTextarea.value
        .split('\n')
        .map(label => label.trim())
        .filter(label => label.length > 0);
    
    if (newLabels.length === 0) {
        alert('Please enter at least one label!');
        return;
    }
    
    PRIORITY_LABELS = newLabels;
    displayPriorityList();
    hideEditLabelsModal();
    
    // Show success message
    alert(`✅ Priority labels updated successfully!\n\nTotal labels: ${PRIORITY_LABELS.length}`);
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
