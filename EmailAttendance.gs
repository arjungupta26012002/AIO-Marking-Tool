const UNFUND_INTERNS_LOG_SHEET_NAME = "Unfound_Interns_Log"; 

function processAttendanceByEmail(data) {
  const { internshipId, sheetName, columnToMark, internNames } = data;

  try {
    const ss = SpreadsheetApp.openById(internshipId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
    }

    const range = sheet.getDataRange();
    const values = range.getValues();
    const header = values[0];
    const internDataRows = values.slice(1);

    const columnToMarkIndex = header.indexOf(columnToMark);
    if (columnToMarkIndex === -1) {
      throw new Error(`Marking column "${columnToMark}" not found.`);
    }

    const EMAIL_COLUMN_INDEX = 3; 

    if (header.length <= EMAIL_COLUMN_INDEX) {
        throw new Error(`Sheet does not have Column D for email lookup.`);
    }

    let foundInternsCount = 0;
    const unfoundInterns = [];

    const lowerCaseProvidedEmails = internNames.map(email => email.toLowerCase().trim());

    for (let i = 0; i < internDataRows.length; i++) {
      const row = internDataRows[i];
      const internEmailInSheet = row[EMAIL_COLUMN_INDEX] ? String(row[EMAIL_COLUMN_INDEX]).toLowerCase().trim() : '';

      if (lowerCaseProvidedEmails.includes(internEmailInSheet)) {
        sheet.getRange(i + 2, columnToMarkIndex + 1).setValue("TRUE");
        foundInternsCount++;
      }
    }

    const foundEmailsInSheet = new Set(
      internDataRows.map(row => (row[EMAIL_COLUMN_INDEX] ? String(row[EMAIL_COLUMN_INDEX]).toLowerCase().trim() : ''))
    );

    lowerCaseProvidedEmails.forEach(providedEmail => {
      if (!foundEmailsInSheet.has(providedEmail)) {
        unfoundInterns.push(providedEmail);
      }
    });

    if (unfoundInterns.length > 0) {
      addUnfoundInterns(internshipId, unfoundInterns, UNFUND_INTERNS_LOG_SHEET_NAME, "Attendance (Email)");
    }

    let message = `Attendance marked by email for ${foundInternsCount} intern(s) in sheet "${sheetName}".`;
    if (unfoundInterns.length > 0) {
      message += ` ${unfoundInterns.length} intern(s) not found by email and added to "${UNFUND_INTERNS_LOG_SHEET_NAME}" sheet.`;
    } else {
      message += ` All provided intern emails were found.`;
    }

    const MARKING_LOG_SPREADSHEET_ID = "MasterID";
    const MARKING_LOG_SHEET_NAME = "logs";
    logMarkingSummary(
      MARKING_LOG_SPREADSHEET_ID,
      MARKING_LOG_SHEET_NAME,
      foundInternsCount,
      Session.getEffectiveUser().getEmail(),
      internshipId
    );

    return message;

  } catch (e) {
    Logger.log(`Error in processAttendanceByEmail for internship ID ${internshipId}, sheet ${sheetName}: ${e.message}`);
    throw new Error(`Failed to mark attendance by email: ${e.message}`);
  }
}
