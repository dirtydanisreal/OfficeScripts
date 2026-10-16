function main(workbook: ExcelScript.Workbook) {
  const addSheet = workbook.getWorksheet("ADD ATTENDING");
  const patientMaker = workbook.getWorksheet("PATIENT_MAKER");

  if (!addSheet || !patientMaker) {
    throw new Error("Required worksheet not found.");
  }

  const name: string = String(addSheet.getRange("B3").getValue())
    .trim()
    .toUpperCase();

  const specialty: string = String(addSheet.getRange("B4").getValue())
    .trim()
    .toUpperCase();

  if (name === "" || specialty === "") {
    return;
  }

  patientMaker.getProtection().pauseProtection("imsorrydave");

  try {
    let targetRangeAddress = "";

    switch (specialty) {
      case "SGE":
      case "SGT":
        targetRangeAddress = "B99:B140";
        break;
      case "SCR":
        targetRangeAddress = "C99:C140";
        break;
      case "SGO":
        targetRangeAddress = "D99:D140";
        break;
      case "URO":
        targetRangeAddress = "E99:E140";
        break;
      case "OMF":
        targetRangeAddress = "F99:F140";
        break;
      case "GYO":
        targetRangeAddress = "G99:G140";
        break;
      case "GYN":
        targetRangeAddress = "H99:H140";
        break;
      case "TXP":
        targetRangeAddress = "I99:I140";
        break;
      case "PLA":
        targetRangeAddress = "J99:J140";
        break;
      case "ENT":
        targetRangeAddress = "K99:K140";
        break;
      case "HM":
        targetRangeAddress = "L99:L140";
        break;
      default:
        throw new Error(`Unsupported specialty: ${specialty}`);
    }

    const range = patientMaker.getRange(targetRangeAddress);
    const values = range.getValues();

    // Check for duplicate
    for (let i = 0; i < values.length; i++) {
      const existing = String(values[i][0] ?? "")
        .trim()
        .toUpperCase();

      if (existing === name) {
        return;
      }
    }

    // Find first blank cell
    for (let i = 0; i < values.length; i++) {
      const existing = String(values[i][0] ?? "").trim();

      if (existing === "") {
        range.getCell(i, 0).setValue(name);
        addSheet.getRange("B3").clear(ExcelScript.ClearApplyTo.contents);
        addSheet.getRange("B4").clear(ExcelScript.ClearApplyTo.contents);
        return;
      }
    }

    throw new Error(`No empty cells available in ${targetRangeAddress}`);
  }
  finally {
    patientMaker.getProtection().resumeProtection();
  }
}
