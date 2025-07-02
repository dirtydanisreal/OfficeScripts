function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  const protection = sheet.getProtection();

  if (sheet.getName() === "PHONE 7.200 ICU") {
    protection.pauseProtection("72icu");
  }
  else{
    protection.pauseProtection("imsorrydave");
  }

  // Define the necessary sheets and ranges
  let listSheet = workbook.getWorksheet("RN-NCT LIST");
  let psSheet = workbook.getWorksheet("PHONE 6.100 AC-PC") || workbook.getWorksheet("PHONE 6.200 AC-PC") || workbook.getWorksheet("PHONE 6.200 ICU") || workbook.getWorksheet("PHONE 7.200 ICU") || workbook.getWorksheet("PHONE 7.100 AC-PC");


  let numbersRange = listSheet.getRange("A1:D900");
  let values = numbersRange.getValues();

  // Find the last used row in the range
  let lastUsedRow = findLastUsedRow(values);

  // Get values from PHONE 6.100 AC-PC sheet (L3, L4, L5)
  const first = psSheet.getRange("L3").getValue().toString().toUpperCase().trim();
  const last = psSheet.getRange("L4").getValue().toString().toUpperCase().trim();
  const num = psSheet.getRange("L5").getValue().toString().toUpperCase().trim();

 if(psSheet.getName() == "PHONE 6.200 ICU" || "PHONE 7.200 ICU")
  {
    const first = psSheet.getRange("Q5").getValue().toString().toUpperCase().trim();
    const last = psSheet.getRange("Q6").getValue().toString().toUpperCase().trim();
    const num = psSheet.getRange("Q7").getValue().toString().toUpperCase().trim();
  }

  // Insert values into RN-NCT LIST sheet at the next available row
  // Note: Using getRangeByIndexes() to set a range with multiple columns
  listSheet.getRangeByIndexes(lastUsedRow, 0, 1, 4).setValues([[first, last, null, num]]);

  // Remove duplicates from columns A, B, and D
  numbersRange.removeDuplicates([0, 1, 3], false);

  if(psSheet.getName() === "PHONE 7.200 ICU" || "PHONE 6.200 ICU")
  {
    psSheet.getRange("Q5:Q7").clear(ExcelScript.ClearApplyTo.contents);
  }

  // Clear contents of L3, L4, and L5 in PHONE 6.100 AC-PC sheet
  psSheet.getRange("L3:L5").clear(ExcelScript.ClearApplyTo.contents);

  // Resume protection
  protection.resumeProtection();
}

// Helper function to find the last used row in the range
function findLastUsedRow(values: (string | number | boolean)[][]): number {
  for (let row = values.length - 1; row >= 0; row--) {
    if (!values[row].every(cell => cell === null || cell === '')) {
      return row + 1; // Return the last used row (1-based index)
    }
  }
  return 0; // Return 0 if no rows are used
}
