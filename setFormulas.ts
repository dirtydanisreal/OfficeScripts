function main(workbook: ExcelScript.Workbook) {
    const sheetNames = ["PHONE 6.100 AC-PC", "PHONE 6.200 AC-PC", "PHONE 7.200 ICU"];

    for (const name of sheetNames) {
        const ps = workbook.getWorksheet(name);
        if (!ps) continue;

        const config = getSheetConfig(name);
        const protection = ps.getProtection();

        if (name === "PHONE 7.200 ICU") {
            protection.pauseProtection("72icu");
        } else {
            protection.pauseProtection("imsorrydave");
        }

        // ✅ Temporarily pause protection (will retain password)
        protection.pauseProtection();

        // ✅ Correct logic for ICU vs AC-PC processing
        if (ps.getName() === "PHONE 7.200 ICU" || ps.getName() === "PHONE 6.200 ICU") {
            applyFormulasICU(ps, config.dataRowStart, config.dataRowEnd);
            applyBorders(ps, `A8:M${config.dataRowEnd}`);
            condFormattingICU(ps, config.dataRowStart, config.dataRowEnd);
            // dataValidationICU(ps, config.dataRowStart, config.dataRowEnd);
            // clearAndSetAllowEdit(ps, config.editableRanges);
        } else {
            applyFormulas(ps, config.dataRowStart, config.dataRowEnd);
            nameAlerts(ps, config.dataRowStart, config.dataRowEnd);
            applyBorders(ps, `A6:I${config.dataRowEnd}`);
            condFormatting(ps, config.dataRowStart, config.dataRowEnd);
            dataValidation(ps, config.dataRowStart, config.dataRowEnd);
            clearAndSetAllowEdit(ps, config.editableRanges);

            try {
                const fontRange = ps.getRange(`A${config.dataRowStart}:I${config.dataRowEnd}`);
                const font = fontRange.getFormat().getFont();
                font.setName("Times New Roman");
                font.setBold(true);
                font.setSize(14);
                fontRange.getFormat().setShrinkToFit(true);
            } catch (err) {
                console.log("Error setting font: " + err.message);
            }

            nameAlerts(ps, config.dataRowStart, config.dataRowEnd);

            makeAllRowsUppercase(ps, config.dataRowStart, config.dataRowEnd); 
        }

        // ✅ Resume protection (will restore original password & settings)
        protection.resumeProtection();
    }
}



function applyFormulas(ps: ExcelScript.Worksheet, startRow: number, endRow: number) {
    const eRange = `E${startRow}:E${endRow}`;
    const gRange = `G${startRow}:G${endRow}`;
    const dRange = `D${startRow}:D${endRow}`;
    const bRange = `B${startRow}:B${endRow}`;
    const cRange = `C${startRow}:C${endRow}`;

   

    const rnFormula = `=SUM(IF(FREQUENCY(IF(${eRange}<>"", MATCH(${eRange}, ${eRange}, 0)), IF(${eRange}<>"", MATCH(${eRange}, ${eRange}, 0))) > 0, 1)) + IF(COUNTIF(${eRange}, C4)=0, 1, 0) + SUM(IF(${gRange} = "RESOURCE",1,0))`;
    const nctFormula = `=SUM(IF(FREQUENCY(IF(${gRange}<>"", MATCH(${gRange}, ${gRange}, 0)), IF(${gRange}<>"", MATCH(${gRange}, ${gRange}, 0))) > 0, 1)) + SUM(IF(${gRange} = "RESOURCE",-1,0))`;
    const acFormula = `=COUNTIFS(${dRange},"*AC*",${bRange},"<>*(ADMIT)*")`;
    const pcFormula = `=COUNTIFS(${dRange},"*PC*",${bRange},"<>*(ADMIT)*")`;

    const otherFormulas = {
        chargeFormula: '=IF(C4<>"",XLOOKUP(C4,\'RN-NCT LIST\'!F:F,\'RN-NCT LIST\'!G:G), "")',
        sgeFormula: `=COUNTIFS(${cRange},"*SGE*",${bRange},"<>*(ADMIT)*")`,
        sgtFormula: `=COUNTIFS(${cRange},"*SGT*",${bRange},"<>*(ADMIT)*")`,
        scrFormula: `=COUNTIFS(${cRange},"*SCR*",${bRange},"<>*(ADMIT)*")`,
        sgoFormula: `=COUNTIFS(${cRange},"*SGO*",${bRange},"<>*(ADMIT)*")`,
        covidFormula: `=COUNTIFS(${bRange},"*(+)*",${bRange},"<>*(ADMIT)*")`,
        sitterFormula: `=SUM(COUNTIFS(${bRange},{"*(SIT)*","*(TS)*","*(72)*"},${bRange},"<>*(ADMIT)*"))`
    };

    // --- Autofill F column ---
    const fStartCell = `F${startRow}`;
    const fStart = ps.getRange(fStartCell);
    fStart.setFormula(`=IF(E${startRow}<>"",XLOOKUP(E${startRow},'RN-NCT LIST'!F:F,'RN-NCT LIST'!G:G),"")`);
    if (endRow > startRow) {
        const fFillRange = ps.getRange(`F${startRow}:F${endRow}`);
        fStart.autoFill(fFillRange, ExcelScript.AutoFillType.fillDefault);
    }

    // --- Autofill H column ---
    const hStartCell = `H${startRow}`;
    const hStart = ps.getRange(hStartCell);
    hStart.setFormula(`=IF(G${startRow}<>"",XLOOKUP(G${startRow},'RN-NCT LIST'!F:F,'RN-NCT LIST'!G:G),"")`);
    if (endRow > startRow) {
        const hFillRange = ps.getRange(`H${startRow}:H${endRow}`);
        hStart.autoFill(hFillRange, ExcelScript.AutoFillType.fillDefault);
    }

    // --- Summary formulas ---
    ps.getRange("O7").setFormula(rnFormula);
    ps.getRange("O8").setFormula(nctFormula);
    ps.getRange("O11").setFormula(acFormula);
    ps.getRange("O12").setFormula(pcFormula);
    ps.getRange("O13").setFormula(`=COUNTIFS(${bRange}, "<>*ADMIT*", ${bRange}, "<>")`);
    ps.getRange("F4").setFormulaLocal(otherFormulas.chargeFormula);
    ps.getRange("O16").setFormulaLocal(otherFormulas.sgeFormula);
    ps.getRange("O17").setFormulaLocal(otherFormulas.sgtFormula);
    ps.getRange("O18").setFormulaLocal(otherFormulas.scrFormula);
    ps.getRange("O19").setFormulaLocal(otherFormulas.sgoFormula);
    ps.getRange("O21").setFormula(`=COUNTIFS(${bRange}, "<>*ADMIT*", ${bRange}, "<>")`);
    ps.getRange("O24").setFormulaLocal(otherFormulas.covidFormula);
    ps.getRange("O27").setFormulaLocal(otherFormulas.sitterFormula);


}


function applyFormulasICU(ps: ExcelScript.Worksheet, startRow: number, endRow: number) {
  const eRange = `G${startRow}:G${endRow}`;
  const gRange = `G${startRow}:G${endRow}`;
  const dRange = `D${startRow}:D${endRow}`;
  //const bRange = `B${startRow}:B${endRow}`;
  const cRange = `C${startRow}:C${endRow}`;



  const rnFormula = `=SUM(IF(FREQUENCY(IF(${eRange}<>"", MATCH(${eRange}, ${eRange}, 0)), IF(${eRange}<>"", MATCH(${eRange}, ${eRange}, 0))) > 0, 1)) + IF(COUNTIF(${eRange}, C4)=0, 1, 0) + SUM(IF(${gRange} = "RESOURCE",1,0))`;
  //const nctFormula = `=SUM(IF(FREQUENCY(IF(${gRange}<>"", MATCH(${gRange}, ${gRange}, 0)), IF(${gRange}<>"", MATCH(${gRange}, ${gRange}, 0))) > 0, 1)) + SUM(IF(${gRange} = "RESOURCE",-1,0))`;
 

  const otherFormulas = {
    chargeFormula: '=IF(C4<>"",XLOOKUP(C4,\'RN-NCT LIST\'!E:E,\'RN-NCT LIST\'!F:F), "")',
    
  };

 

  // --- Autofill H column ---
  const hStartCell = `I${startRow}`;
  const hStart = ps.getRange(hStartCell);
  hStart.setFormula(`=IF(G${startRow}<>"",XLOOKUP(G${startRow},'RN-NCT LIST'!E:E,'RN-NCT LIST'!F:F),"")`);
  if (endRow > startRow) {
    const hFillRange = ps.getRange(`I${startRow}:I${endRow}`);
    hStart.autoFill(hFillRange, ExcelScript.AutoFillType.fillDefault);
  }

  // --- Summary formulas ---
  ps.getRange("O7").setFormula(rnFormula);
 
  ps.getRange("F4").setFormulaLocal(otherFormulas.chargeFormula);
 
 // ps.getRange("O21").setFormula(`=COUNTIFS(${bRange}, "<>*ADMIT*", ${bRange}, "<>")`);

    const cells = ["I10", "I12", "I14", "I16", "I18", "I20"];
    for (let cell of cells) {
        ps.getRange(cell).getFormat().getFill().setColor("FFFFFF");
    }
 
}

function makeAllRowsUppercase(sheet: ExcelScript.Worksheet, startRow: number, endRow: number) {
    const range = sheet.getRangeByIndexes(startRow - 1, 0, endRow - startRow + 1, sheet.getUsedRange().getColumnCount());
    const values = range.getValues();
    const uppercased = values.map(row =>
        row.map(cell => typeof cell === "string" ? cell.toUpperCase() : cell)
    );
    range.setValues(uppercased);
}


function applyBorders(sheet: ExcelScript.Worksheet, rangeAddress: string) {
    let borders = [
        ExcelScript.BorderIndex.edgeTop,
        ExcelScript.BorderIndex.edgeRight,
        ExcelScript.BorderIndex.edgeLeft,
        ExcelScript.BorderIndex.edgeBottom,
        ExcelScript.BorderIndex.insideHorizontal,
        ExcelScript.BorderIndex.insideVertical
    ];

    let weights = [ExcelScript.BorderWeight.thick, ExcelScript.BorderWeight.thick, ExcelScript.BorderWeight.thick,
    ExcelScript.BorderWeight.thick, ExcelScript.BorderWeight.thin, ExcelScript.BorderWeight.thin];

    borders.forEach((edge, index) => {
        const border = sheet.getRange(rangeAddress).getFormat().getRangeBorder(edge);
        border.setStyle(ExcelScript.BorderLineStyle.continuous);
        border.setWeight(weights[index]);
        border.setColor("000000");
    });
}


function condFormatting(sheet: ExcelScript.Worksheet, startRow: number, endRow: number) {
    const patientName = sheet.getRange(`B${startRow}:B${endRow}`);

    try {
        const formats = patientName.getConditionalFormats();
        for (let i = formats.length - 1; i >= 0; i--) {
            formats[i].delete();
        }
    } catch (error) {
        console.log("Error clearing formats: " + error.message);
    }

    applyConditionalFormatting(patientName, '(S)', 'red', '00b0f0');
    applyConditionalFormatting(patientName, '(P)', 'red', '00b0f0');

    const formulaRule = patientName.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
    formulaRule.getCustom().getRule().setFormula(`=AND(M${startRow}>1, M${endRow}<5)`);

    let format = formulaRule.getCustom().getFormat();
    format.getFill().setColor("33ccff");
    format.getFont().setBold(true);
}

function condFormattingICU(sheet: ExcelScript.Worksheet, startRow: number, endRow: number) {
  const patientName = sheet.getRange(`D${startRow}:D${endRow}`);

  try {
    const formats = patientName.getConditionalFormats();
    for (let i = formats.length - 1; i >= 0; i--) {
      formats[i].delete();
    }
  } catch (error) {
    console.log("Error clearing formats: " + error.message);
  }

  applyConditionalFormatting(patientName, '(S)', 'red', '00b0f0');
  applyConditionalFormatting(patientName, '(P)', 'red', '00b0f0');

  const formulaRule = patientName.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
  formulaRule.getCustom().getRule().setFormula(`=AND(N${startRow}>1, N${endRow}<5)`);

  let format = formulaRule.getCustom().getFormat();
  format.getFill().setColor("33ccff");
  format.getFont().setBold(true);
}



// Helper function for conditional formatting
function applyConditionalFormatting(range: ExcelScript.Range, text: string, color1: string, color2: string) {
    const textComparison = range.addConditionalFormat(ExcelScript.ConditionalFormatType.containsText).getTextComparison();
    textComparison.setRule({ text, operator: ExcelScript.ConditionalTextOperator.contains });
    textComparison.getFormat().getFill().setColor(color1);
    textComparison.getFormat().getFont().setBold(true);
}


function dataValidation(sheet: ExcelScript.Worksheet, startRow: number, endRow: number) {
    const rnRange = sheet.getRange(`E${startRow}:E${endRow}`);
    const nctRange = sheet.getRange(`G${startRow}:G${endRow}`);
    const locRange = sheet.getRange(`D${startRow}:D${endRow}`);
    const teamRange = sheet.getRange(`C${startRow}:C${endRow}`);
    const shiftRange = sheet.getRange("C5");

    [rnRange, nctRange, locRange, teamRange, shiftRange].forEach(r => r.getDataValidation()?.clear());

    const listRule = (src: string): ExcelScript.DataValidationRule => ({ list: { source: src, inCellDropDown: true } });
    const errorAlert = (msg: string): ExcelScript.DataValidationErrorAlert => ({
        message: msg, showAlert: false, style: ExcelScript.DataValidationAlertStyle.warning, title: msg
    });

    rnRange.getDataValidation()?.setRule(listRule("='RN-NCT LIST'!$F:$F"));
    rnRange.getDataValidation()?.setIgnoreBlanks(true);
    rnRange.getDataValidation()?.setErrorAlert(errorAlert("rn not found"));

    nctRange.getDataValidation()?.setRule(listRule("='RN-NCT LIST'!$F:$F"));
    nctRange.getDataValidation()?.setIgnoreBlanks(true);
    nctRange.getDataValidation()?.setErrorAlert(errorAlert("nct not found"));

    locRange.getDataValidation()?.setRule(listRule("AC,PC"));
    locRange.getDataValidation()?.setIgnoreBlanks(true);
    locRange.getDataValidation()?.setErrorAlert(errorAlert("loc doesn't exist"));

    teamRange.getDataValidation()?.setRule(listRule("=$B$78:$B$128"));
    teamRange.getDataValidation()?.setIgnoreBlanks(true);
    teamRange.getDataValidation()?.setErrorAlert(errorAlert("wrong team"));

    shiftRange.getDataValidation()?.setRule(listRule("7 AM,7 PM"));
    shiftRange.getDataValidation()?.setErrorAlert(errorAlert("not valid shift"));
}

function dataValidationICU(sheet: ExcelScript.Worksheet, startRow: number, endRow: number) {
  const rnRange = sheet.getRange(`G${startRow}:G${endRow}`);
 
 
  const teamRange = sheet.getRange(`C${startRow}:C${endRow}`);
  const shiftRange = sheet.getRange("D3");

  [rnRange, teamRange, shiftRange].forEach(r => r.getDataValidation()?.clear());

  const listRule = (src: string): ExcelScript.DataValidationRule => ({ list: { source: src, inCellDropDown: true } });
  const errorAlert = (msg: string): ExcelScript.DataValidationErrorAlert => ({
    message: msg, showAlert: false, style: ExcelScript.DataValidationAlertStyle.warning, title: msg
  });

  rnRange.getDataValidation()?.setRule(listRule("='RN-NCT LIST'!$F:$F"));
  rnRange.getDataValidation()?.setIgnoreBlanks(true);
  rnRange.getDataValidation()?.setErrorAlert(errorAlert("rn not found"));

  teamRange.getDataValidation()?.setRule(listRule("=$B$78:$B$128"));
  teamRange.getDataValidation()?.setIgnoreBlanks(true);
  teamRange.getDataValidation()?.setErrorAlert(errorAlert("wrong team"));

  shiftRange.getDataValidation()?.setRule(listRule("7 AM,7 PM"));
  shiftRange.getDataValidation()?.setErrorAlert(errorAlert("not valid shift"));
}







function allowEdit(sheet: ExcelScript.Worksheet) {
    const config = getSheetConfig(sheet.getName());
    const protection = sheet.getProtection();

    if(sheet.getName() === "PHONE 7.200 ICU"){
        if (protection.getProtected()) protection.pauseProtectionForAllAllowEditRanges("72icu");
    }
    else {
        if (protection.getProtected()) protection.pauseProtectionForAllAllowEditRanges("imsorrydave");
    }

    // Lock all cells first (default)
    sheet.getUsedRange().getFormat().getProtection().setLocked(true);

    // Unlock only editable ranges from config
    for (const rangeStr of config.editableRanges) {
        try {
            const range = sheet.getRange(rangeStr);
            range.getFormat().getProtection().setLocked(false);
        } catch (err) {
            console.log(`Could not unlock range ${rangeStr}: ${err.message}`);
        }
    }

    // Re-apply protection after locking/unlocking
   
    protection.protect();
    console.log("Protection reapplied with config-defined editable ranges.");
}

/**
 * Clears all existing allowToEdit settings and then sets new allowToEdit ranges.
 *
 * @param workbook The workbook object in the Excel script.
 * @param editableRanges An array of string ranges that should be editable, e.g., ["A1:C10", "D1:D10"]
 */
function clearAndSetAllowEdit(sheet: ExcelScript.Worksheet, editableRanges: string[]) {
    const protection = sheet.getProtection();

    // Fully unprotect the sheet




    // Add new editable ranges
    for (let i = 0; i < editableRanges.length; i++) {
        try {
            protection.addAllowEditRange(`EditRange${i + 1}`, editableRanges[i], { password: "imsorrydave" });
        } catch (err) {
            console.log(`Error adding allow edit range '${editableRanges[i]}': ${err.message}`);
        }
    }
}






function nameAlerts(sheet: ExcelScript.Worksheet, startRow: number, endRow: number) {
    const range = sheet.getRange(`B${startRow}:B${endRow}`);
    const data = range.getValues();
    const numRows = data.length;
    const columnMValues: (number | string)[][] = [];

    const excluded = ["(S)", "(P)", "(SIT)", "(ADMIT)"];
    const redFill = ["(S)", "(P)"];
    const blueFill = ["(SIT)", "(ADMIT)"];
    const redRGB = "#FF0000", blueRGB = "#0000FF";

    function clean(name: string): string {
        name = name.split(",")[0];
        excluded.forEach(text => {
            name = name.replace(text, "");
        });
        return name.trim().toLowerCase().replace(/[^a-z]/g, "");
    }

    function normalize(name: string): string {
        return name.replace(/(.)\1+/g, "$1");
    }

    const names = data.map(row => row[0] ? clean(row[0] as string) : "");
    const normNames = names.map(name => normalize(name));

    const counts: { [key: string]: number } = {};
    normNames.forEach(name => { if (name) counts[name] = (counts[name] || 0) + 1; });

    for (let i = 0; i < numRows; i++) {
        const fullName = data[i][0] as string;
        const cell = sheet.getRange(`B${startRow + i}`);
        const fill = cell.getFormat().getFill().getColor();

        if (!fullName) {
            columnMValues.push([0]);
            continue;
        }

        const cleaned = clean(fullName);
        const normalized = normalize(cleaned);
        const words = cleaned.split(" ");
        const hasDupWords = new Set(words).size !== words.length;
        const isMultiple = counts[normalized] > 1;

        columnMValues.push([isMultiple && !hasDupWords ? 2 : 1]);

        if (fill === redRGB && !redFill.some(t => fullName.includes(t))) cell.getFormat().getFill().setColor("#FFFFFF");
        if (fill === blueRGB && !blueFill.some(t => fullName.includes(t))) cell.getFormat().getFill().setColor("#FFFFFF");
    }

    sheet.getRangeByIndexes(startRow - 1, 12, numRows, 1).setValues(columnMValues);
}




function getSheetConfig(sheetName: string): { range: string, editableRanges: string[], dataRowStart: number, dataRowEnd: number } {
    if (sheetName === "PHONE 6.100 AC-PC") {
        return {
            range: "A7:I38",
            editableRanges: ["C4", "C5", "B7:B38", "C7:C38", "D7:D38", "E7:E38", "G7:G38", "L3:L5"],
            dataRowStart: 7,
            dataRowEnd: 38
        };
    } else if (sheetName === "PHONE 6.200 AC-PC") {
        return {
            range: "A7:I26",
            editableRanges: ["C4", "C5", "B7:B26", "C7:C26", "D7:D26", "E7:E26", "G7:G26", "L3:L5"],
            dataRowStart: 7,
            dataRowEnd: 26
        };
    } else if (sheetName === "PHONE 7.200 ICU") {
        return {
            range: "A9:M20",
            editableRanges: ["C4", "C5", "C9:C20", "D9:D20", "G9:G20", "Q5:Q7", "J9:J20", "M9:M20"],
            dataRowStart: 9,
            dataRowEnd: 20
        };
    } else if (sheetName === "PHONE 6.200 ICU") {
        return {
            range: "A7:I18",
            editableRanges: ["C4", "C5", "B7:B18", "C7:C18", "D7:D18", "G7:G18","Q5:Q7"],
            dataRowStart: 7,
            dataRowEnd: 18
        };
    } else {
        throw new Error(`Unsupported sheet: ${sheetName}`);
    }
}

