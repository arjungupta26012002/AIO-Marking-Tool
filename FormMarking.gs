function processFormMarking(data) {

  if (!data.internshipId || !data.sheetName || !data.columnToMark || !data.internNames || data.internNames.length === 0) {
    throw new Error("Missing required data for form marking process.");
  }

  return processAttendanceByEmail({
    internshipId: data.internshipId,
    sheetName: data.sheetName,
    columnToMark: data.columnToMark,
    internNames: data.internNames, 
    processType: "Form Marking" 
  });
}
