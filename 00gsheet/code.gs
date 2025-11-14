// No need to hardcode sheet ID anymore - we'll pass it from the frontend
function getSheetById(sheetId) {
  if (!sheetId) {
    throw new Error('Sheet ID is required');
  }
  return SpreadsheetApp.openById(sheetId).getActiveSheet();
}

function doPost(e) {
  return handlePostRequest(e);
}

function handlePostRequest(e) {
  var action = e.parameter.action;
  var sheetId = e.parameter.sheetId; // Get sheetId from request
  var result = {};
  
  try {
    // Check if user is authenticated
    var userEmail = Session.getActiveUser().getEmail();
    Logger.log('User accessing: ' + userEmail);
    
    switch(action) {
      case 'getData':
        result = getSheetData(sheetId);
        break;
      case 'saveData':
        var rows = JSON.parse(e.parameter.rows);
        result = saveData(sheetId, rows);
        break;
      case 'addRow':
        var rowData = JSON.parse(e.parameter.rowData);
        result = addRow(sheetId, rowData);
        break;
      case 'deleteRow':
        var rowId = e.parameter.rowId;
        result = deleteRow(sheetId, rowId);
        break;
      case 'testConnection':
        result = testConnection(sheetId);
        break;
      default:
        result = { success: false, message: 'Unknown action' };
    }
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    result = { success: false, message: error.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'POST')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// Handle CORS preflight requests
function doOptions() {
  return ContentService.createTextOutput()
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type')
    .setHeader('Access-Control-Max-Age', '86400');
}

function getSheetData(sheetId) {
  try {
    var sheet = getSheetById(sheetId);
    var data = sheet.getDataRange().getValues();
    
    if (data.length === 0) {
      return { success: false, message: 'Sheet is empty' };
    }
    
    var headers = data[0];
    var result = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        row[headers[j]] = data[i][j];
      }
      row.id = i;
      result.push(row);
    }
    
    return {
      success: true,
      headers: headers,
      data: result,
      sheetName: sheet.getName()
    };
  } catch (error) {
    return { success: false, message: 'Error accessing sheet: ' + error.toString() };
  }
}

function saveData(sheetId, rows) {
  try {
    var sheet = getSheetById(sheetId);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var rowNum = parseInt(row.id) + 1;
      
      if (rowNum <= sheet.getLastRow()) {
        var rowData = [];
        for (var j = 0; j < headers.length; j++) {
          rowData.push(row[headers[j]] || '');
        }
        sheet.getRange(rowNum, 1, 1, headers.length).setValues([rowData]);
      }
    }
    
    return { success: true, message: 'Data saved successfully' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function addRow(sheetId, rowData) {
  try {
    var sheet = getSheetById(sheetId);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    var newRow = [];
    for (var i = 0; i < headers.length; i++) {
      newRow.push(rowData[headers[i]] || '');
    }
    
    sheet.appendRow(newRow);
    
    return { success: true, message: 'Row added successfully', newRow: newRow };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function deleteRow(sheetId, rowId) {
  try {
    var sheet = getSheetById(sheetId);
    var rowNum = parseInt(rowId) + 1;
    
    if (rowNum <= sheet.getLastRow()) {
      sheet.deleteRow(rowNum);
      return { success: true, message: 'Row deleted successfully' };
    } else {
      return { success: false, message: 'Row not found' };
    }
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function testConnection(sheetId) {
  try {
    var sheet = getSheetById(sheetId);
    var data = sheet.getDataRange().getValues();
    return { 
      success: true, 
      message: 'Connection successful', 
      sheetName: sheet.getName(),
      rowCount: data.length,
      columnCount: data.length > 0 ? data[0].length : 0
    };
  } catch (error) {
    return { success: false, message: 'Connection failed: ' + error.toString() };
  }
}