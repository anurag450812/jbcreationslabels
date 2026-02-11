// ============================================
// COMPLETE GOOGLE APPS SCRIPT
// Copy this entire code to your Google Apps Script
// ============================================

// ============================================
// YOUR EXISTING FUNCTIONS (updated to read threshold from S7)
// ============================================

function findDaysForNeedInSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("STOCK ANALYSIS");

  if (!sheet) {
    Logger.log("Error: Sheet 'STOCK ANALYSIS' not found.");
    return;
  }

  const needCell = sheet.getRange("S3");
  const daysCell = sheet.getRange("S11");
  const thresholdCell = sheet.getRange("S7");
  const targetNeedThreshold = thresholdCell.getValue() || 150; // Read from S7, default to 150 if empty

  let initialNeedValue = needCell.getValue();
  let currentDays = daysCell.getValue();

  Logger.log("--- findDaysForNeedInSheet Start ---");
  Logger.log("Sheet Name: " + sheet.getName());
  Logger.log("Value in S3 (needCell): " + initialNeedValue);
  Logger.log("Type of S3 value: " + typeof initialNeedValue);
  Logger.log("isNaN(initialNeedValue): " + isNaN(initialNeedValue));
  Logger.log("Value in S11 (daysCell): " + currentDays);
  Logger.log("Type of S11 value: " + typeof currentDays);
  Logger.log("isNaN(currentDays): " + isNaN(currentDays));
  Logger.log("Value in S7 (targetNeedThreshold): " + targetNeedThreshold);

  if (typeof initialNeedValue !== 'number' || isNaN(initialNeedValue)) {
    Logger.log("Error: S3 does not contain a valid number.");
    return;
  }
  if (typeof currentDays !== 'number' || isNaN(currentDays)) {
    currentDays = 0;
    daysCell.setValue(currentDays);
    initialNeedValue = needCell.getValue();
    Logger.log("Days was not a valid number. Set to 0. New S3 value: " + initialNeedValue);
  }
  Logger.log("---------------------------------------");

  if (initialNeedValue > targetNeedThreshold) {
    Logger.log("Need is too high (" + initialNeedValue + "). Decreasing days.");

    while (true) {
      currentDays = currentDays - 1;

      if (currentDays < 0) {
        daysCell.setValue(0);
        let finalNeed = needCell.getValue();
        Logger.log("Days reduced to 0. Need is now: " + finalNeed);
        return;
      }

      daysCell.setValue(currentDays);
      let newNeed = needCell.getValue();

      if (newNeed <= targetNeedThreshold) {
        daysCell.setValue(currentDays + 1);
        break;
      }
    }
  }
  else {
    Logger.log("Need is too low (" + initialNeedValue + "). Increasing days.");

    if (currentDays < 0) {
      currentDays = 0;
      daysCell.setValue(currentDays);
    }

    const MAX_DAYS_LIMIT = 25;

    while (true) {
      currentDays = currentDays + 1;

      if (currentDays > MAX_DAYS_LIMIT) {
        daysCell.setValue(MAX_DAYS_LIMIT);
        let finalNeed = needCell.getValue();
        Logger.log("Days at max limit (" + MAX_DAYS_LIMIT + "). Need is now: " + finalNeed);
        return;
      }

      daysCell.setValue(currentDays);
      let newNeed = needCell.getValue();

      if (newNeed > targetNeedThreshold) {
        break;
      }
    }
  }

  let finalDays = daysCell.getValue();
  let finalNeed = needCell.getValue();

  Logger.log("Days adjusted to: " + finalDays + ", Need: " + finalNeed);
  Logger.log("--- findDaysForNeedInSheet End ---");
}

function onMySheetEdit(e) {
  const spreadsheet = e.source;
  const sheet = spreadsheet.getActiveSheet();
  const editedRange = e.range;

  const targetSpreadsheetName = "order wali sheet";
  const targetSheetName = "STOCK COUNT";
  const targetColumnIndex = 2; // Column B

  if (spreadsheet.getName() !== targetSpreadsheetName) {
    Logger.log("Edit not in target spreadsheet: " + spreadsheet.getName());
    return;
  }

  if (sheet.getName() !== targetSheetName) {
    Logger.log("Edit not in target sheet: " + sheet.getName());
    return;
  }

  if (editedRange.getColumn() === targetColumnIndex) {
    Logger.log("Edit in column B of " + targetSheetName + ". Calling findDaysForNeedInSheet.");
    findDaysForNeedInSheet();
  } else {
    Logger.log("Edit not in target column B.");
  }
}

// ============================================
// NEW FUNCTIONS FOR JB CREATIONS WEBSITE
// Stock Deduction from Label Processing
// ============================================

/**
 * Web-safe version of findDaysForNeedInSheet
 * This version works when called from the website (doPost)
 * It does NOT use SpreadsheetApp.getUi().alert() which fails in web context
 */
function findDaysForNeedInSheet_WebSafe() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("STOCK ANALYSIS");

  if (!sheet) {
    Logger.log("Error: Sheet 'STOCK ANALYSIS' not found.");
    return { success: false, message: "Sheet 'STOCK ANALYSIS' not found" };
  }

  const needCell = sheet.getRange("S3");
  const daysCell = sheet.getRange("S11");
  const thresholdCell = sheet.getRange("S7");
  const targetNeedThreshold = thresholdCell.getValue() || 150; // Read from S7, default to 150 if empty

  let initialNeedValue = needCell.getValue();
  let currentDays = daysCell.getValue();

  Logger.log("--- findDaysForNeedInSheet_WebSafe Start ---");
  Logger.log("Sheet Name: " + sheet.getName());
  Logger.log("Value in S3 (needCell): " + initialNeedValue);
  Logger.log("Value in S11 (daysCell): " + currentDays);
  Logger.log("Value in S7 (targetNeedThreshold): " + targetNeedThreshold);

  if (typeof initialNeedValue !== 'number' || isNaN(initialNeedValue)) {
    Logger.log("Error: S3 is not a valid number");
    return { success: false, message: "S3 is not a valid number" };
  }
  
  if (typeof currentDays !== 'number' || isNaN(currentDays)) {
    currentDays = 0;
    daysCell.setValue(currentDays);
    initialNeedValue = needCell.getValue();
    Logger.log("Days was not a valid number. Set to 0.");
  }

  // Case 1: Current Need is above the target threshold
  if (initialNeedValue > targetNeedThreshold) {
    Logger.log("Need is too high (" + initialNeedValue + "). Decreasing days.");

    while (true) {
      currentDays = currentDays - 1;

      if (currentDays < 0) {
        daysCell.setValue(0);
        let finalNeed = needCell.getValue();
        Logger.log("Days reduced to 0. Need is now: " + finalNeed);
        return { 
          success: true, 
          message: "Days reduced to 0", 
          days: 0, 
          need: finalNeed 
        };
      }

      daysCell.setValue(currentDays);
      SpreadsheetApp.flush(); // Force the sheet to recalculate
      let newNeed = needCell.getValue();

      if (newNeed <= targetNeedThreshold) {
        daysCell.setValue(currentDays + 1);
        SpreadsheetApp.flush();
        break;
      }
    }
  }
  // Case 2: Current Need is at or below the target threshold
  else {
    Logger.log("Need is too low (" + initialNeedValue + "). Increasing days.");

    if (currentDays < 0) {
      currentDays = 0;
      daysCell.setValue(currentDays);
    }

    const MAX_DAYS_LIMIT = 25;

    while (true) {
      currentDays = currentDays + 1;

      if (currentDays > MAX_DAYS_LIMIT) {
        daysCell.setValue(MAX_DAYS_LIMIT);
        SpreadsheetApp.flush();
        let finalNeed = needCell.getValue();
        Logger.log("Days increased to limit (" + MAX_DAYS_LIMIT + "). Need is now: " + finalNeed);
        return { 
          success: true, 
          message: "Days at max limit", 
          days: MAX_DAYS_LIMIT, 
          need: finalNeed 
        };
      }

      daysCell.setValue(currentDays);
      SpreadsheetApp.flush(); // Force the sheet to recalculate
      let newNeed = needCell.getValue();

      if (newNeed > targetNeedThreshold) {
        break;
      }
    }
  }

  let finalDays = daysCell.getValue();
  let finalNeed = needCell.getValue();

  Logger.log("Days adjusted to: " + finalDays + ", Need: " + finalNeed);
  Logger.log("--- findDaysForNeedInSheet_WebSafe End ---");
  
  return { 
    success: true, 
    message: "Days adjusted successfully", 
    days: finalDays, 
    need: finalNeed 
  };
}

/**
 * Handles POST requests from the JB Creations website
 * This function handles stock deduction and priority labels management
 */
function doPost(e) {
  try {
    // Parse the incoming data
    var data;
    
    // Handle different content types
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.data) {
      data = JSON.parse(e.parameter.data);
    } else {
      throw new Error('No data received');
    }
    
    Logger.log("--- doPost Request Received ---");
    Logger.log("Action: " + data.action);
    Logger.log("Data: " + JSON.stringify(data));
    
    if (data.action === 'deductStock') {
      return handleDeductStock(data);
    }
    
    if (data.action === 'savePriorityLabels') {
      return handleSavePriorityLabels(data);
    }
    
    if (data.action === 'saveLabelCriteria') {
      return handleSaveLabelCriteria(data);
    }
    
    if (data.action === 'saveAppConfig') {
      return handleSaveAppConfig(data);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Unknown action: ' + data.action
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log("Error in doPost: " + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handles the stock deduction logic
 * @param {Object} data - Contains sheetName and labelCounts
 */
function handleDeductStock(data) {
  var sheetName = data.sheetName || 'STOCK COUNT'; // Default to your STOCK COUNT sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log("Sheet not found: " + sheetName);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Sheet not found: ' + sheetName
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  var labelCounts = data.labelCounts;
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  var updates = [];
  var errors = [];
  
  // Process each label
  for (var label in labelCounts) {
    var deductAmount = labelCounts[label];
    var found = false;
    
    // Search for the label in column A (case-insensitive, partial match)
    // This will find 'bd18' in 'jb-BD18', 'JB-bd18-xl', etc.
    for (var i = 1; i < values.length; i++) { // Start from 1 to skip header row
      var cellLabel = String(values[i][0]).toLowerCase().trim();
      var searchLabel = label.toLowerCase().trim();
      
      // Use includes() for partial matching instead of exact match
      // This finds the label anywhere within the cell value
      if (cellLabel.includes(searchLabel)) {
        found = true;
        var currentStock = Number(values[i][1]) || 0;
        var newStock = Math.max(0, currentStock - deductAmount); // Don't go below 0
        
        // Update the cell (Column B = index 2, row = i + 1 because rows are 1-indexed)
        sheet.getRange(i + 1, 2).setValue(newStock);
        
        updates.push({
          label: label,
          matchedCell: values[i][0], // Show what cell was matched
          previousStock: currentStock,
          deducted: deductAmount,
          newStock: newStock
        });
        
        Logger.log("Updated " + label + " (matched '" + values[i][0] + "'): " + currentStock + " -> " + newStock + " (deducted " + deductAmount + ")");
        break;
      }
    }
    
    if (!found) {
      errors.push({
        label: label,
        message: 'Label not found in sheet (searched for partial match)'
      });
      Logger.log("Label not found: " + label);
    }
  }
  
  // Trigger the existing findDaysForNeedInSheet function after stock update
  // Use the web-safe version that doesn't use UI alerts
  var daysAdjustmentResult = null;
  try {
    daysAdjustmentResult = findDaysForNeedInSheet_WebSafe();
    Logger.log("findDaysForNeedInSheet_WebSafe triggered after stock update");
    Logger.log("Days adjustment result: " + JSON.stringify(daysAdjustmentResult));
  } catch (triggerError) {
    Logger.log("Error triggering findDaysForNeedInSheet_WebSafe: " + triggerError.toString());
    daysAdjustmentResult = { success: false, message: triggerError.toString() };
  }
  
  // Write the deducted labels to "Orders From Stock" sheet
  var ordersResult = null;
  try {
    ordersResult = writeOrdersFromStock(labelCounts);
    Logger.log("Orders From Stock updated: " + JSON.stringify(ordersResult));
  } catch (ordersError) {
    Logger.log("Error writing to Orders From Stock: " + ordersError.toString());
    ordersResult = { success: false, message: ordersError.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: 'Stock updated successfully',
    updates: updates,
    errors: errors,
    totalUpdated: updates.length,
    totalErrors: errors.length,
    daysAdjustment: daysAdjustmentResult,
    ordersFromStock: ordersResult
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handles GET requests - useful for testing if the web app is working
 * Also handles getPriorityLabels and getLabelCriteria actions
 */
function doGet(e) {
  var action = e.parameter ? e.parameter.action : null;
  
  // Handle getPriorityLabels action
  if (action === 'getPriorityLabels') {
    return handleGetPriorityLabels();
  }
  
  // Handle getLabelCriteria action
  if (action === 'getLabelCriteria') {
    return handleGetLabelCriteria();
  }
  
  // Handle getAppConfig action
  if (action === 'getAppConfig') {
    return handleGetAppConfig();
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    message: 'JB Creations Stock Update API is running',
    timestamp: new Date().toISOString(),
    availableActions: ['deductStock', 'savePriorityLabels', 'getPriorityLabels', 'saveLabelCriteria', 'getLabelCriteria', 'saveAppConfig', 'getAppConfig']
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// PRIORITY LABELS MANAGEMENT
// Store priority labels in a dedicated sheet
// ============================================

var PRIORITY_LABELS_SHEET_NAME = 'PRIORITY_LABELS';

/**
 * Get or create the priority labels sheet
 */
function getPriorityLabelsSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(PRIORITY_LABELS_SHEET_NAME);
  
  if (!sheet) {
    // Create the sheet if it doesn't exist
    sheet = spreadsheet.insertSheet(PRIORITY_LABELS_SHEET_NAME);
    
    // Add header
    sheet.getRange('A1').setValue('Priority Label SKU');
    sheet.getRange('A1').setFontWeight('bold');
    sheet.getRange('A1').setBackground('#4a90d9');
    sheet.getRange('A1').setFontColor('white');
    sheet.setColumnWidth(1, 200);
    
    Logger.log('Created new PRIORITY_LABELS sheet');
  }
  
  return sheet;
}

/**
 * Handle saving priority labels to Google Sheets
 */
function handleSavePriorityLabels(data) {
  try {
    var labels = data.labels;
    
    if (!labels || !Array.isArray(labels) || labels.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'No labels provided'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var sheet = getPriorityLabelsSheet();
    
    // Clear existing labels (except header)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 1).clear();
    }
    
    // Write new labels starting from row 2
    var labelData = labels.map(function(label) {
      return [label];
    });
    
    sheet.getRange(2, 1, labelData.length, 1).setValues(labelData);
    
    // Add timestamp in column B, row 1
    sheet.getRange('B1').setValue('Last Updated');
    sheet.getRange('B1').setFontWeight('bold');
    sheet.getRange('B2').setValue(new Date().toISOString());
    sheet.setColumnWidth(2, 180);
    
    Logger.log('Saved ' + labels.length + ' priority labels to sheet');
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Priority labels saved successfully',
      count: labels.length,
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('Error saving priority labels: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error saving labels: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle getting priority labels from Google Sheets
 */
function handleGetPriorityLabels() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(PRIORITY_LABELS_SHEET_NAME);
    
    if (!sheet) {
      Logger.log('Priority labels sheet not found, returning empty array');
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        labels: [],
        message: 'Priority labels sheet not found'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      // Only header exists, no labels
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        labels: [],
        message: 'No priority labels found'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Get labels from column A, starting from row 2
    var range = sheet.getRange(2, 1, lastRow - 1, 1);
    var values = range.getValues();
    
    // Flatten and filter empty values
    var labels = values
      .map(function(row) { return String(row[0]).trim(); })
      .filter(function(label) { return label.length > 0; });
    
    // Get last updated timestamp if available
    var lastUpdated = null;
    try {
      lastUpdated = sheet.getRange('B2').getValue();
    } catch (e) {}
    
    Logger.log('Retrieved ' + labels.length + ' priority labels from sheet');
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      labels: labels,
      count: labels.length,
      lastUpdated: lastUpdated ? lastUpdated.toString() : null,
      message: 'Priority labels retrieved successfully'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('Error getting priority labels: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      labels: [],
      message: 'Error getting labels: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Test function for priority labels
 */
function testPriorityLabels() {
  // Test saving
  var testSaveData = {
    action: 'savePriorityLabels',
    labels: ['test-label-1', 'test-label-2', 'bd18', 'bd21']
  };
  
  var saveResult = handleSavePriorityLabels(testSaveData);
  Logger.log('Save Test Result: ' + saveResult.getContent());
  
  // Test getting
  var getResult = handleGetPriorityLabels();
  Logger.log('Get Test Result: ' + getResult.getContent());
}

// ============================================
// LABEL DETECTION CRITERIA MANAGEMENT
// Store label detection criteria in a dedicated sheet
// ============================================

var LABEL_CRITERIA_SHEET_NAME = 'LABEL_CRITERIA';

/**
 * Get or create the label criteria sheet
 */
function getLabelCriteriaSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(LABEL_CRITERIA_SHEET_NAME);
  
  if (!sheet) {
    // Create the sheet if it doesn't exist
    sheet = spreadsheet.insertSheet(LABEL_CRITERIA_SHEET_NAME);
    
    // Add headers
    sheet.getRange('A1').setValue('Setting Key');
    sheet.getRange('B1').setValue('Setting Value');
    sheet.getRange('A1:B1').setFontWeight('bold');
    sheet.getRange('A1:B1').setBackground('#4a90d9');
    sheet.getRange('A1:B1').setFontColor('white');
    sheet.setColumnWidth(1, 250);
    sheet.setColumnWidth(2, 400);
    
    Logger.log('Created new LABEL_CRITERIA sheet');
  }
  
  return sheet;
}

/**
 * Handle saving label detection criteria to Google Sheets
 */
function handleSaveLabelCriteria(data) {
  try {
    var criteria = data.criteria;
    
    if (!criteria) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'No criteria provided'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var sheet = getLabelCriteriaSheet();
    
    // Clear existing criteria (except header)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 2).clear();
    }
    
    // Write criteria as key-value pairs starting from row 2
    var criteriaData = [
      ['meesho.returnAddressPatternString', criteria.meesho.returnAddressPatternString || ''],
      ['meesho.returnAddressFlags', criteria.meesho.returnAddressFlags || 'i'],
      ['meesho.couriers', JSON.stringify(criteria.meesho.couriers || [])],
      ['meesho.pickupKeyword', criteria.meesho.pickupKeyword || 'PICKUP'],
      ['flipkart.shippingAddressKeyword', criteria.flipkart.shippingAddressKeyword || ''],
      ['flipkart.soldByPatternString', criteria.flipkart.soldByPatternString || ''],
      ['flipkart.soldByFlags', criteria.flipkart.soldByFlags || 'i']
    ];
    
    sheet.getRange(2, 1, criteriaData.length, 2).setValues(criteriaData);
    
    // Add timestamp
    sheet.getRange('C1').setValue('Last Updated');
    sheet.getRange('C1').setFontWeight('bold');
    sheet.getRange('C2').setValue(new Date().toISOString());
    sheet.setColumnWidth(3, 180);
    
    Logger.log('Saved label detection criteria to sheet');
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Label detection criteria saved successfully',
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('Error saving label criteria: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error saving criteria: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle getting label detection criteria from Google Sheets
 */
function handleGetLabelCriteria() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(LABEL_CRITERIA_SHEET_NAME);
    
    if (!sheet) {
      Logger.log('Label criteria sheet not found, returning null');
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        criteria: null,
        message: 'Label criteria sheet not found - using defaults'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      // Only header exists, no criteria
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        criteria: null,
        message: 'No criteria found - using defaults'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Get all criteria key-value pairs
    var range = sheet.getRange(2, 1, lastRow - 1, 2);
    var values = range.getValues();
    
    // Build criteria object
    var criteria = {
      meesho: {},
      flipkart: {}
    };
    
    for (var i = 0; i < values.length; i++) {
      var key = values[i][0];
      var value = values[i][1];
      
      if (key && value) {
        var parts = key.split('.');
        if (parts.length === 2) {
          var platform = parts[0];
          var setting = parts[1];
          
          // Parse JSON arrays
          if (setting === 'couriers') {
            try {
              value = JSON.parse(value);
            } catch (e) {
              Logger.log('Error parsing couriers JSON: ' + e);
            }
          }
          
          criteria[platform][setting] = value;
        }
      }
    }
    
    // Get last updated timestamp if available
    var lastUpdated = null;
    try {
      lastUpdated = sheet.getRange('C2').getValue();
    } catch (e) {}
    
    Logger.log('Retrieved label detection criteria from sheet');
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      criteria: criteria,
      lastUpdated: lastUpdated ? lastUpdated.toString() : null,
      message: 'Label detection criteria retrieved successfully'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('Error getting label criteria: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      criteria: null,
      message: 'Error getting criteria: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Test function for label criteria
 */
function testLabelCriteria() {
  // Test saving
  var testSaveData = {
    action: 'saveLabelCriteria',
    criteria: {
      meesho: {
        returnAddressPatternString: 'if\\s+undelivered,?\\s+return\\s+to:?\\s*([A-Za-z0-9_-]+)',
        returnAddressFlags: 'i',
        couriers: ['DELHIVERY', 'SHADOWFAX', 'VALMO', 'XPRESS BEES'],
        pickupKeyword: 'PICKUP'
      },
      flipkart: {
        shippingAddressKeyword: 'Shipping/Customer address:',
        soldByPatternString: 'Sold\\s+By:?\\s*([A-Za-z0-9\\s_-]+?)(?:\\s*,|\\s*\\n|\\s*$|\\s+[A-Z])',
        soldByFlags: 'i'
      }
    }
  };
  
  var saveResult = handleSaveLabelCriteria(testSaveData);
  Logger.log('Save Test Result: ' + saveResult.getContent());
  
  // Test getting
  var getResult = handleGetLabelCriteria();
  Logger.log('Get Test Result: ' + getResult.getContent());
}

// ============================================
// APP CONFIGURATION MANAGEMENT
// Store app configuration (Google Sheets settings) in a dedicated sheet
// ============================================

var APP_CONFIG_SHEET_NAME = 'APP_CONFIG';

/**
 * Get or create the app configuration sheet
 */
function getAppConfigSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(APP_CONFIG_SHEET_NAME);
  
  if (!sheet) {
    // Create the sheet if it doesn't exist
    sheet = spreadsheet.insertSheet(APP_CONFIG_SHEET_NAME);
    
    // Add headers
    sheet.getRange('A1').setValue('Setting Key');
    sheet.getRange('B1').setValue('Setting Value');
    sheet.getRange('A1:B1').setFontWeight('bold');
    sheet.getRange('A1:B1').setBackground('#10b981');
    sheet.getRange('A1:B1').setFontColor('white');
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 400);
    
    Logger.log('Created new APP_CONFIG sheet');
  }
  
  return sheet;
}

/**
 * Handle saving app configuration to Google Sheets
 */
function handleSaveAppConfig(data) {
  try {
    var config = data.config;
    
    if (!config) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'No configuration provided'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var sheet = getAppConfigSheet();
    
    // Clear existing config (except header)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 2).clear();
    }
    
    // Write config as key-value pairs starting from row 2
    var configData = [
      ['spreadsheetId', config.spreadsheetId || ''],
      ['sheetName', config.sheetName || 'STOCK COUNT'],
      ['apiKey', config.apiKey || ''],
      ['webAppUrl', config.webAppUrl || '']
    ];
    
    sheet.getRange(2, 1, configData.length, 2).setValues(configData);
    
    // Add timestamp
    sheet.getRange('C1').setValue('Last Updated');
    sheet.getRange('C1').setFontWeight('bold');
    sheet.getRange('C2').setValue(new Date().toISOString());
    sheet.setColumnWidth(3, 180);
    
    Logger.log('Saved app configuration to sheet');
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'App configuration saved successfully',
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('Error saving app config: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error saving config: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle getting app configuration from Google Sheets
 */
function handleGetAppConfig() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(APP_CONFIG_SHEET_NAME);
    
    if (!sheet) {
      Logger.log('App config sheet not found, returning empty config');
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        config: {},
        message: 'App config sheet not found'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      // Only header exists, no config
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        config: {},
        message: 'No configuration found'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Get all config key-value pairs
    var range = sheet.getRange(2, 1, lastRow - 1, 2);
    var values = range.getValues();
    
    // Build config object
    var config = {};
    
    for (var i = 0; i < values.length; i++) {
      var key = values[i][0];
      var value = values[i][1];
      
      if (key) {
        config[key] = value || '';
      }
    }
    
    // Get last updated timestamp if available
    var lastUpdated = null;
    try {
      lastUpdated = sheet.getRange('C2').getValue();
    } catch (e) {}
    
    Logger.log('Retrieved app configuration from sheet');
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      config: config,
      lastUpdated: lastUpdated ? lastUpdated.toString() : null,
      message: 'App configuration retrieved successfully'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('Error getting app config: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      config: {},
      message: 'Error getting config: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Test function for app configuration
 */
function testAppConfig() {
  // Test saving
  var testSaveData = {
    action: 'saveAppConfig',
    config: {
      spreadsheetId: '1234567890abcdef',
      sheetName: 'STOCK COUNT',
      apiKey: 'test-api-key',
      webAppUrl: 'https://script.google.com/macros/s/test/exec'
    }
  };
  
  var saveResult = handleSaveAppConfig(testSaveData);
  Logger.log('Save Test Result: ' + saveResult.getContent());
  
  // Test getting
  var getResult = handleGetAppConfig();
  Logger.log('Get Test Result: ' + getResult.getContent());
}

/**
 * Test function to manually test the stock deduction
 * You can run this from the Apps Script editor to test
 */
function testDeductStock() {
  var testData = {
    action: 'deductStock',
    sheetName: 'STOCK COUNT',
    labelCounts: {
      'bd18': 2,
      'bd21': 1
    }
  };
  
  // Simulate the request
  var mockEvent = {
    postData: {
      contents: JSON.stringify(testData)
    }
  };
  
  var result = doPost(mockEvent);
  Logger.log("Test Result: " + result.getContent());
}

/**
 * Get current stock for a label (utility function)
 * Uses partial matching - finds 'bd18' in 'jb-BD18', etc.
 */
function getStockForLabel(labelName, sheetName) {
  sheetName = sheetName || 'STOCK COUNT';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if (!sheet) return null;
  
  var values = sheet.getDataRange().getValues();
  var searchLabel = labelName.toLowerCase().trim();
  
  for (var i = 1; i < values.length; i++) {
    var cellLabel = String(values[i][0]).toLowerCase().trim();
    // Use includes() for partial matching
    if (cellLabel.includes(searchLabel)) {
      return {
        label: values[i][0],
        stock: values[i][1],
        row: i + 1,
        matchedSearch: labelName
      };
    }
  }
  
  return null;
}

// ============================================
// ORDERS FROM STOCK TRACKING FUNCTIONS
// These functions write deducted labels to "Orders From Stock" sheet
// and manage data archiving to "Orders From Stock yesterday"
// ============================================

/**
 * Writes deducted labels to the "Orders From Stock" sheet
 * Each label is written as many times as its count
 * For example: bd18 with count 15 will write "bd18" 15 times in column A
 * Date is stored in column B for tracking
 * 
 * @param {Object} labelCounts - Object with label names as keys and counts as values
 */
function writeOrdersFromStock(labelCounts) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get or create the "Orders From Stock" sheet
  var ordersSheet = spreadsheet.getSheetByName("Orders From Stock");
  if (!ordersSheet) {
    ordersSheet = spreadsheet.insertSheet("Orders From Stock");
    Logger.log("Created new sheet: Orders From Stock");
  }
  
  // First, archive old data (move yesterday's data, clear older data)
  try {
    archiveOrdersData();
  } catch (archiveError) {
    Logger.log("Warning: Error during archive, continuing with write: " + archiveError.toString());
  }
  
  // Current date for tracking when data was entered
  var currentDate = new Date();
  var dateString = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  // Prepare rows to append
  var rowsToAppend = [];
  
  for (var label in labelCounts) {
    var count = labelCounts[label];
    // Write the label 'count' number of times
    for (var i = 0; i < count; i++) {
      rowsToAppend.push([label, dateString]);
    }
  }
  
  if (rowsToAppend.length > 0) {
    // Re-fetch the sheet reference after archive operations
    ordersSheet = spreadsheet.getSheetByName("Orders From Stock");
    if (!ordersSheet) {
      ordersSheet = spreadsheet.insertSheet("Orders From Stock");
    }
    
    // Find the last row with data
    var lastRow = ordersSheet.getLastRow();
    
    // Append all rows at once for efficiency
    var startRow = lastRow + 1;
    ordersSheet.getRange(startRow, 1, rowsToAppend.length, 2).setValues(rowsToAppend);
    SpreadsheetApp.flush(); // Force write to complete
    
    Logger.log("Added " + rowsToAppend.length + " rows to Orders From Stock with date " + dateString);
  }
  
  return {
    success: true,
    rowsAdded: rowsToAppend.length,
    date: dateString
  };
}

/**
 * Helper function to convert any date value to yyyy-MM-dd string
 * Handles Date objects, date strings, and various formats
 * 
 * @param {*} dateValue - The date value to convert
 * @returns {string} Date in yyyy-MM-dd format, or empty string if invalid
 */
function toDateString(dateValue) {
  if (!dateValue) return '';
  
  // Handle Date objects
  if (dateValue instanceof Date) {
    return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  
  // Handle string values
  var str = String(dateValue).trim();
  
  // Already in yyyy-MM-dd format
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    return str;
  }
  
  // Try to parse as date
  try {
    var parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
  } catch (e) {}
  
  return '';
}

/**
 * Archives orders data:
 * 1. Moves yesterday's data from "Orders From Stock" to "Orders From Stock yesterday"
 * 2. Clears data older than yesterday from "Orders From Stock yesterday"
 */
function archiveOrdersData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the sheets
  var ordersSheet = spreadsheet.getSheetByName("Orders From Stock");
  var yesterdaySheet = spreadsheet.getSheetByName("Orders From Stock yesterday");
  
  // Create "Orders From Stock yesterday" if it doesn't exist
  if (!yesterdaySheet) {
    yesterdaySheet = spreadsheet.insertSheet("Orders From Stock yesterday");
    Logger.log("Created new sheet: Orders From Stock yesterday");
  }
  
  if (!ordersSheet) {
    Logger.log("Orders From Stock sheet not found, skipping archive");
    return;
  }
  
  // Calculate today and yesterday dates
  var today = new Date();
  var todayString = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  var yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  var yesterdayString = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  Logger.log("Archive - Today: " + todayString + ", Yesterday: " + yesterdayString);
  
  // First, clear old data from "Orders From Stock yesterday" (older than yesterday)
  clearOldDataFromYesterdaySheet(yesterdaySheet, yesterdayString);
  
  // Then, move yesterday's data from "Orders From Stock" to "Orders From Stock yesterday"
  moveYesterdayDataToArchive(ordersSheet, yesterdaySheet, todayString, yesterdayString);
}

/**
 * Clears data older than yesterday from the "Orders From Stock yesterday" sheet
 * 
 * @param {Sheet} yesterdaySheet - The "Orders From Stock yesterday" sheet
 * @param {string} yesterdayString - Yesterday's date in yyyy-MM-dd format
 */
function clearOldDataFromYesterdaySheet(yesterdaySheet, yesterdayString) {
  var lastRow = yesterdaySheet.getLastRow();
  
  if (lastRow === 0) {
    Logger.log("Orders From Stock yesterday sheet is empty, nothing to clear");
    return;
  }
  
  // Get all data from the sheet
  var data = yesterdaySheet.getRange(1, 1, lastRow, 2).getValues();
  
  // Find rows to keep (only rows with yesterday's date)
  var rowsToKeep = [];
  var rowsCleared = 0;
  
  for (var i = 0; i < data.length; i++) {
    var rowDate = toDateString(data[i][1]);
    
    // Keep only rows from yesterday
    if (rowDate === yesterdayString) {
      rowsToKeep.push(data[i]);
    } else if (rowDate && rowDate !== "") {
      rowsCleared++;
    }
  }
  
  // Clear the sheet and rewrite only the rows to keep
  if (rowsCleared > 0) {
    yesterdaySheet.clear();
    
    if (rowsToKeep.length > 0) {
      yesterdaySheet.getRange(1, 1, rowsToKeep.length, 2).setValues(rowsToKeep);
    }
    
    SpreadsheetApp.flush();
    Logger.log("Cleared " + rowsCleared + " old rows from Orders From Stock yesterday, kept " + rowsToKeep.length);
  } else {
    Logger.log("No old data to clear from Orders From Stock yesterday");
  }
}

/**
 * Moves yesterday's data from "Orders From Stock" to "Orders From Stock yesterday"
 * Keeps only today's data in "Orders From Stock"
 * 
 * @param {Sheet} ordersSheet - The "Orders From Stock" sheet
 * @param {Sheet} yesterdaySheet - The "Orders From Stock yesterday" sheet
 * @param {string} todayString - Today's date in yyyy-MM-dd format
 * @param {string} yesterdayString - Yesterday's date in yyyy-MM-dd format
 */
function moveYesterdayDataToArchive(ordersSheet, yesterdaySheet, todayString, yesterdayString) {
  var lastRow = ordersSheet.getLastRow();
  
  if (lastRow === 0) {
    Logger.log("Orders From Stock sheet is empty, nothing to move");
    return;
  }
  
  // Get all data from the Orders From Stock sheet
  var data = ordersSheet.getRange(1, 1, lastRow, 2).getValues();
  
  // Separate data by date
  var yesterdayData = [];
  var todaysData = [];
  var olderData = [];
  
  for (var i = 0; i < data.length; i++) {
    var rowDate = toDateString(data[i][1]);
    
    if (rowDate === todayString) {
      // Keep today's data in the main sheet
      todaysData.push(data[i]);
    } else if (rowDate === yesterdayString) {
      // Move yesterday's data to archive
      yesterdayData.push(data[i]);
    } else if (data[i][0] && String(data[i][0]).trim() !== "") {
      // Data older than yesterday or without proper date - also move to archive
      olderData.push(data[i]);
      Logger.log("Found older data: label=" + data[i][0] + ", date=" + rowDate);
    }
  }
  
  // Move yesterday's data (and any older data) to the archive sheet
  var dataToArchive = yesterdayData.concat(olderData);
  if (dataToArchive.length > 0) {
    var lastRowInYesterday = yesterdaySheet.getLastRow();
    var startRow = lastRowInYesterday + 1;
    yesterdaySheet.getRange(startRow, 1, dataToArchive.length, 2).setValues(dataToArchive);
    Logger.log("Moved " + dataToArchive.length + " rows to Orders From Stock yesterday (" + 
               yesterdayData.length + " from yesterday, " + olderData.length + " older)");
  } else {
    Logger.log("No yesterday/old data to move from Orders From Stock");
  }
  
  // Only rewrite if we actually moved something (avoid unnecessary clear)
  if (dataToArchive.length > 0) {
    // Rewrite the Orders From Stock sheet with only today's data
    ordersSheet.clear();
    
    if (todaysData.length > 0) {
      ordersSheet.getRange(1, 1, todaysData.length, 2).setValues(todaysData);
    }
    
    SpreadsheetApp.flush();
    Logger.log("Orders From Stock now contains " + todaysData.length + " rows (today's data only)");
  } else {
    Logger.log("No data needed to be moved, Orders From Stock unchanged");
  }
}

/**
 * Test function to manually test the orders tracking
 */
function testWriteOrdersFromStock() {
  var testLabelCounts = {
    'bd18': 5,
    'bd21': 3,
    'ml03': 2
  };
  
  var result = writeOrdersFromStock(testLabelCounts);
  Logger.log("Test Result: " + JSON.stringify(result));
}

/**
 * Test function to manually run the archive process
 */
function testArchiveOrdersData() {
  archiveOrdersData();
  Logger.log("Archive process completed");
}
