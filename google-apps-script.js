// ============================================
// COMPLETE GOOGLE APPS SCRIPT
// Copy this entire code to your Google Apps Script
// ============================================

// ============================================
// YOUR EXISTING FUNCTIONS (unchanged)
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
  const targetNeedThreshold = 350;

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

    const MAX_DAYS_LIMIT = 450;

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
  const targetNeedThreshold = 350;

  let initialNeedValue = needCell.getValue();
  let currentDays = daysCell.getValue();

  Logger.log("--- findDaysForNeedInSheet_WebSafe Start ---");
  Logger.log("Sheet Name: " + sheet.getName());
  Logger.log("Value in S3 (needCell): " + initialNeedValue);
  Logger.log("Value in S11 (daysCell): " + currentDays);

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

    const MAX_DAYS_LIMIT = 450;

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
 * This function deducts stock counts based on processed labels
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
  
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: 'Stock updated successfully',
    updates: updates,
    errors: errors,
    totalUpdated: updates.length,
    totalErrors: errors.length,
    daysAdjustment: daysAdjustmentResult
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handles GET requests - useful for testing if the web app is working
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    message: 'JB Creations Stock Update API is running',
    timestamp: new Date().toISOString(),
    availableActions: ['deductStock']
  })).setMimeType(ContentService.MimeType.JSON);
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
