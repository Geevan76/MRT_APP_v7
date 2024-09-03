// Adjustable variables
var maxImageWidth = 100;  // Maximum image width in pixels
var maxImageHeight = 100; // Maximum image height in pixels

// Variable to hold the Google Doc ID for the template
var templateId = '118GMPtbOBvrXtLyhc3bnwO1XojVCPZsTiaIGkv9Z6uk'; // Replace with your Google Doc ID

// Adjustable column mapping
var columnMapping = {
  '{{Inspection ID}}': 2,  // Column B (index 1)
  '{{UserName}}': 5,       // Column E (index 4)
  '{{trainNo}}': 7,        // Column G (index 6)
  '{{Location}}': 8,       // Column H (index 7)
  '{{Car Body}}': 11,      // Column K (index 10)
  '{{Section Name}}': 13,  // Column M (index 12)
  '{{Subsystem Name}}': 15,// Column O (index 14)
  '{{Serial Number}}': 16, // Column P (index 15)
  '{{Subcomponent}}': 18,  // Column R (index 17)
  '{{Condition}}': 19,     // Column S (index 18)
  '{{Defect Type}}': 20,   // Column T (index 19)
  '{{Remarks}}': 21,       // Column U (index 20)
  '{{Image URL}}': 27,     // Column AA (index 26),
  '{{Image ID}}': 23       // Column W (index 22)
};

function generateFunctionalReport() {
  generateReport('Functional_Inspection_Report', 'F');
}

function generateVisualReport() {
  generateReport('Visual_Inspection_Report', 'V');
}

function generateReport(sheetName, prefix) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var filter = sheet.getFilter(); // Get the active filter applied to the sheet
  var trainNumbers = [];
  var visibleData = [];

  if (filter) {
    // Loop through the rows starting from row 10 to the last row
    var lastRow = sheet.getLastRow();
    for (var row = 10; row <= lastRow; row++) {
      var trainNo = sheet.getRange(row, 7).getValue().trim(); // Trim whitespace

      // Check if the row is visible and has a non-empty train number
      if (trainNo !== "" && sheet.isRowHiddenByFilter(row) === false) {
        if (trainNumbers.indexOf(trainNo) === -1) {
          trainNumbers.push(trainNo);
        }
        visibleData.push(sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]); // Collect visible row data
      }

      // If more than one unique train number is found, raise an error
      if (trainNumbers.length > 1) {
        SpreadsheetApp.getUi().alert('Error: More than one train number is present in the filtered data. Please ensure only one train number is present.');
        return;
      }
    }
  }

  if (trainNumbers.length === 0) {
    SpreadsheetApp.getUi().alert('Error: No train number found. Please check your data.');
    return;
  }

  // Continue with report generation using the single unique train number
  var trainNo = trainNumbers[0];

  var doc = createDocumentFromTemplate(templateId, trainNo, sheetName);
  
  replaceTrainNoPlaceholder(doc, trainNo, prefix); // Replace the {{trainNo}} placeholder with the unique train number
  
  replaceAllPlaceholders(doc); // Replace any remaining placeholders with empty strings

  var body = doc.getBody();
  removePlaceholderRow(body); // Remove the placeholder row before populating the data

  // Populate the report with only the visible, filtered data
  populateTemplateTableWithData(body, visibleData);

  saveDocumentToFolder(doc, sheetName, trainNo);

  updateSheetWithReportDetails(sheet, getFileName(sheetName, trainNo), doc);

  SpreadsheetApp.getUi().alert('Report generated and saved for Train No: ' + trainNo);
}

function createDocumentFromTemplate(templateId, trainNo, sheetName) {
  var templateDoc = DriveApp.getFileById(templateId);
  var fileName = getFileName(sheetName, trainNo);
  var copyDoc = templateDoc.makeCopy(fileName);
  var doc = DocumentApp.openById(copyDoc.getId());

  // Set the document to A4 landscape mode
  setDocumentToLandscape(doc);

  return doc;
}

function setDocumentToLandscape(doc) {
  var body = doc.getBody();
  
  // Set page size for A4 in landscape mode (595 points height x 842 points width)
  body.setPageHeight(595).setPageWidth(842);

  // Optionally adjust margins if needed
  body.setMarginTop(36);      // 36 points = 0.5 inches
  body.setMarginBottom(36);   // 36 points = 0.5 inches
  body.setMarginLeft(36);     // 36 points = 0.5 inches
  body.setMarginRight(36);    // 36 points = 0.5 inches
}

function replaceTrainNoPlaceholder(doc, trainNo, prefix) {
  var suffix = prefix === 'F' ? 'Functional' : 'Visual';
  var fullTrainNoText = trainNo + ' (' + suffix + ')';

  // Replace the placeholder in the document body
  var body = doc.getBody();
  body.replaceText('{{trainNo}}', fullTrainNoText);

  // Replace the placeholder in the header
  var header = doc.getHeader();
  if (header) {
    header.replaceText('{{trainNo}}', fullTrainNoText);
  }
}

function replaceAllPlaceholders(doc) {
  var placeholders = Object.keys(columnMapping);
  var body = doc.getBody();

  placeholders.forEach(function(placeholder) {
    body.replaceText(placeholder, '');
  });
}

function removePlaceholderRow(body) {
  // Assuming there's only one table in the body now
  var tables = body.getTables();
  var table = tables[0]; // Retrieve the first and only table
 
  // Ensure the table has more than one row
  if (table.getNumRows() > 1) {
    var secondRow = table.getRow(1);
    table.removeRow(1); // Remove the second row (index 1)
  }
}

function populateTemplateTableWithData(body, data) {
  // Get all tables in the document body
  var tables = body.getTables();

  if (tables.length === 0) {
    throw new Error("No tables found in the document body. Ensure the template has a table for data.");
  }

  var table = tables[0]; // Retrieve the first and only table

  // Loop through the data and append rows to the table
  data.forEach(function(row, index) {
    var tableRow = table.appendTableRow();
    tableRow.appendTableCell((index + 1).toString()); // Running number starting from 1

    tableRow.appendTableCell(row[columnMapping['{{Location}}'] - 1].toString()); // Location
    tableRow.appendTableCell(row[columnMapping['{{Car Body}}'] - 1].toString()); // Car Body
    tableRow.appendTableCell(row[columnMapping['{{UserName}}'] - 1].toString()); // User Name
    tableRow.appendTableCell(row[columnMapping['{{Section Name}}'] - 1].toString()); // Section Name
    tableRow.appendTableCell(row[columnMapping['{{Subsystem Name}}'] - 1].toString()); // Subsystem Name
    tableRow.appendTableCell(row[columnMapping['{{Serial Number}}'] - 1].toString()); // Serial Number
    tableRow.appendTableCell(row[columnMapping['{{Subcomponent}}'] - 1].toString()); // Subcomponent
    tableRow.appendTableCell(row[columnMapping['{{Condition}}'] - 1].toString()); // Condition
    tableRow.appendTableCell(row[columnMapping['{{Defect Type}}'] - 1].toString()); // Defect Type
    tableRow.appendTableCell(row[columnMapping['{{Remarks}}'] - 1].toString()); // Remarks

    var imageUrl = row[columnMapping['{{Image URL}}'] - 1]; // Image URL
    var imageCell = tableRow.appendTableCell();

    if (imageUrl) {
      try {
        var response = UrlFetchApp.fetch(imageUrl);
        var imageBlob = response.getBlob();

        if (imageBlob.getContentType().startsWith('image/')) {
          // Apply the user-defined max width and height to the image
          imageCell.appendImage(imageBlob).setWidth(maxImageWidth).setHeight(maxImageHeight);
        } else {
          imageCell.appendParagraph('Invalid image type');
        }
      } catch (e) {
        imageCell.appendParagraph('Failed to load image: ' + e.toString());
      }
    } else {
      imageCell.appendParagraph('No image available');
    }
  });
}

function saveDocumentToFolder(doc, sheetName, trainNo) {
  var spreadsheetFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parentFolder = spreadsheetFile.getParents().next();

  var reportFolderName = sheetName + 's'; // Functional_Inspection_Reports or Visual_Inspection_Reports
  var reportFolder = getOrCreateFolder(parentFolder, reportFolderName);

  var trainFolder = getOrCreateFolder(reportFolder, trainNo);

  var file = DriveApp.getFileById(doc.getId());
  file.moveTo(trainFolder);
}

function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

function getFileName(sheetName, trainNo) {
  return trainNo + ' ' + (sheetName === 'Functional_Inspection_Report' ? 'Functional' : 'Visual') + ' Inspection Report';
}

function updateSheetWithReportDetails(sheet, fileName, doc) {
  sheet.getRange('H6').setValue(fileName); // Report Name
  sheet.getRange('I6').setValue(doc.getUrl()); // Report URL

  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  sheet.getRange('J6').setValue(timestamp); // Timestamp
}
