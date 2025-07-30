function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

function getInternshipList() {
  const masterSpreadsheetId = "ID-of-sheet";
  const refSheetName = "RefID";

  try {
    const ss = SpreadsheetApp.openById(masterSpreadsheetId);
    const masterSheet = ss.getSheetByName(refSheetName);

    if (!masterSheet) {
      throw new Error(`Sheet "${refSheetName}" not found in the master spreadsheet (ID: ${masterSpreadsheetId}).`);
    }

    const range = masterSheet.getDataRange();
    const values = range.getValues();

    const internshipData = values.slice(1).map(row => ({
      name: row[0],
      id: row[1]
    }));
    return internshipData;
  } catch (e) {
    Logger.log(`Error in getInternshipList: ${e.message}`);
    throw new Error(`Could not fetch internship list: ${e.message}. Please check the master spreadsheet ID and sheet name.`);
  }
}

function getSheetNames(spreadsheetId) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheets = ss.getSheets();
    return sheets.map(sheet => sheet.getName());
  } catch (e) {
    Logger.log(`Error getting sheet names for ID ${spreadsheetId}: ${e.message}`);
    throw new Error(`Could not open spreadsheet with ID: ${spreadsheetId}. Please ensure the ID is correct and you have access.`);
  }
}

function getColumnNames(spreadsheetId, sheetName) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found in spreadsheet ID: ${spreadsheetId}.`);
    }
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    return headerRow;
  } catch (e) {
    Logger.log(`Error getting column names for sheet ${sheetName} in ID ${spreadsheetId}: ${e.message}`);
    throw new Error(`Could not retrieve column names. Please check the sheet name and ID.`);
  }
}

function addUnfoundInterns(targetSpreadsheetId, internNames, logSheetName, processType) {
  try {
    const ss = SpreadsheetApp.openById(targetSpreadsheetId);
    let logSheet = ss.getSheetByName(logSheetName);

    if (!logSheet) {
      logSheet = ss.insertSheet(logSheetName);
      logSheet.getRange(1, 1).setValue("Timestamp");
      logSheet.getRange(1, 2).setValue("Unfound Intern Name/Email");
      logSheet.getRange(1, 3).setValue("Process Type");
      logSheet.setFrozenRows(1);
      logSheet.autoResizeColumns(1, 3);
    }

    const dataToWrite = internNames.map(name => [new Date(), name, processType]);
    let lastRow = logSheet.getLastRow();

    if (lastRow > 0) {
      logSheet.insertRowsAfter(lastRow, 1);
      lastRow++;
    }

    logSheet.getRange(lastRow + 1, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
    logSheet.insertRowsAfter(lastRow + dataToWrite.length, 1);

    Logger.log(`Added ${internNames.length} unfound intern(s) to "${logSheetName}" in spreadsheet ID: ${targetSpreadsheetId}.`);
  } catch (e) {
    Logger.log(`Error adding unfound interns to spreadsheet ID ${targetSpreadsheetId}: ${e.message}`);
    throw new Error(`Failed to log unfound interns: ${e.message}`);
  }
}

function logMarkingSummary(logSpreadsheetId, logSheetName, markedCount, userEmail, loggedInternshipId) { 
  try {
    const logSs = SpreadsheetApp.openById(logSpreadsheetId);
    let logSheet = logSs.getSheetByName(logSheetName);

    if (!logSheet) {
      logSheet = logSs.insertSheet(logSheetName);

      logSheet.getRange(1, 1).setValue("Date");
      logSheet.getRange(1, 2).setValue("User Email");
      logSheet.getRange(1, 3).setValue("Interns Marked Count");
      logSheet.getRange(1, 4).setValue("Internship ID"); 
      logSheet.setFrozenRows(1);
      logSheet.autoResizeColumns(1, 4); 
    }

    logSheet.appendRow([
      new Date(),
      userEmail,
      markedCount,
      loggedInternshipId 
    ]);
  } catch (e) {
    Logger.log(`Error logging marking summary: ${e.message}`);
  }
}
