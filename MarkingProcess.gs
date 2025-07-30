function processInternMarking(data) {
  const { internshipId, sheetName, columnToMark, internNames: originalInternNames, markerValue, processType } = data;

  const FULL_NAME_COLUMN_INDEX = 2; 
  const UNFUND_INTERNS_LOG_SHEET_NAME = "Unfound_Interns_Log";
  const MARKING_LOG_SPREADSHEET_ID = "MasterSheet-ID";
  const MARKING_LOG_SHEET_NAME = "logs";

  if (!internshipId || !sheetName || !columnToMark || !originalInternNames || originalInternNames.length === 0 || !markerValue || !processType) {
    throw new Error("Missing required data for marking process.");
  }

  let processedInternNames = [];
  if (processType && processType.includes("Deliverables")) {
    processedInternNames = originalInternNames.map(name => {
      let cleanedName = String(name).trim();

      const pathMatch = cleanedName.match(/.*[\\/]([^\\/]+)$/);
      if (pathMatch && pathMatch[1]) {
        cleanedName = pathMatch[1].trim();
      }

      const underscoreIndex = cleanedName.indexOf('_');
      if (underscoreIndex !== -1) {
        cleanedName = cleanedName.substring(0, underscoreIndex);
      }
      return cleanedName.trim();
    });
    Logger.log(`Cleaned names for "${processType}": ${processedInternNames}`);
  } else {
    processedInternNames = originalInternNames.map(name => String(name).trim());
    Logger.log(`Using original (trimmed) names for "${processType}": ${processedInternNames}`);
  }

  try {
    const ss = SpreadsheetApp.openById(internshipId);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found in the selected internship spreadsheet.`);
    }

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnToMarkIndex = headerRow.indexOf(columnToMark);

    if (columnToMarkIndex === -1) {
      throw new Error(`Column "${columnToMark}" not found in sheet "${sheetName}".`);
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const sheetInternNames = values.slice(1).map(row => row[FULL_NAME_COLUMN_INDEX]);

    const foundInterns = [];
    const unfoundInterns = [];
    const updates = [];

    const internMap = new Map();
    sheetInternNames.forEach((name, index) => {
      if (name) {
        internMap.set(String(name).toLowerCase().trim(), index + 2);
      }
    });

    processedInternNames.forEach(submittedName => {
      const submittedNameLower = String(submittedName).toLowerCase().trim();
      let found = false;

      if (internMap.has(submittedNameLower)) {
        const row = internMap.get(submittedNameLower);
        updates.push({ row: row, col: columnToMarkIndex + 1, value: markerValue });
        foundInterns.push(submittedName);
        found = true;
      }

      if (!found) {
        unfoundInterns.push(submittedName);
      }
    });

    if (updates.length > 0) {
      updates.forEach(update => {
        sheet.getRange(update.row, update.col).setValue(update.value);
      });
      Logger.log(`Successfully marked ${updates.length} intern(s) in sheet "${sheetName}".`);
    }

    if (unfoundInterns.length > 0) {
      addUnfoundInterns(internshipId, unfoundInterns, UNFUND_INTERNS_LOG_SHEET_NAME, processType);
    }

    logMarkingSummary(
      MARKING_LOG_SPREADSHEET_ID,
      MARKING_LOG_SHEET_NAME,
      foundInterns.length,
      Session.getEffectiveUser().getEmail(),
      internshipId
    );

    let message = `${processType} marked for ${foundInterns.length} intern(s) in sheet "${sheetName}".`;
    if (unfoundInterns.length > 0) {
      message += ` ${unfoundInterns.length} intern(s) not found and added to "${UNFUND_INTERNS_LOG_SHEET_NAME}" sheet in the same spreadsheet.`;
    } else {
      message += ` All provided intern names were found.`;
    }
    return message;

  } catch (e) {
    Logger.log(`Error in processInternMarking for ${processType}: ${e.message}`);
    throw new Error(`Failed to process ${processType}: ${e.message}. Please review inputs and permissions.`);
  }
}

function processAttendance(data) {
  return processInternMarking({
    ...data,
    markerValue: "TRUE",
    processType: "Attendance"
  });
}

function processWeek1Deliverables(data) {
  return processInternMarking({
    ...data,
    markerValue: "TRUE",
    processType: "Weekly Deliverables"
  });
}
