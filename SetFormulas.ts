function main(workbook: ExcelScript.Workbook) {
  const sheetNames = [
    "PHONE 6.100 AC-PC", "PHONE 6.200 AC-PC", "PHONE 7.200 ICU",
    "PHONE 10.200 AC-PC", "PHONE 6.100 UNI", "PHONE 6.100 ONC",
    "PHONE 6.200 ICU", "PHONE 7.200 UNI"
  ];

  createValidationNamedRanges(workbook);

  for (const name of sheetNames) {
    const sheet = workbook.getWorksheet(name);
    if (!sheet) continue;

    const config = getSheetConfig(name);
    const protection = sheet.getProtection();
    const wasProtected = protection.getProtected();
    const password = getProtectionPassword(name);
    const isICU = name.toUpperCase().includes("ICU");
    const state = ensureEditableForWrites(workbook, sheet, password);

    if (!state.ok) {
      console.log(`Could not pause or remove protection on "${name}". Sheet skipped.`);
      continue;
    }

    if (isICU) {
      applyFormulasICU(sheet, config.dataRowStart, config.dataRowEnd);
      writeRNCount(sheet, config.dataRowStart, config.dataRowEnd);
      writeNCTCount(sheet, config.dataRowStart, config.dataRowEnd);
      applyBorders(sheet, `A8:M${config.dataRowEnd}`);
      condFormattingICU(sheet, config.dataRowStart, config.dataRowEnd);
      dataValidationICU(sheet, config.dataRowStart, config.dataRowEnd);
    } else {
      applyFormulas(sheet, config.dataRowStart, config.dataRowEnd);
      nameAlerts(sheet, config.dataRowStart, config.dataRowEnd);
      writeRNCount(sheet, config.dataRowStart, config.dataRowEnd);
      writeNCTCount(sheet, config.dataRowStart, config.dataRowEnd);
      applyBorders(sheet, `A6:I${config.dataRowEnd}`);
      condFormatting(sheet, config.dataRowStart, config.dataRowEnd);
      dataValidation(sheet, config.dataRowStart, config.dataRowEnd);

      const editable = [...config.editableRanges];
      if (name === "PHONE 10.200 AC-PC") {
        const qCells = ["Q14", "Q19", "Q24", "Q29"];
        for (const address of qCells) {
          const range = sheet.getRange(address);
          range.getFormat().getProtection().setLocked(false);
          range.getFormat().getFill().setColor("#d7e1eb");
          range.getFormat().getFont().setBold(true);
        }
        editable.push(...qCells);
      }
      clearAndSetAllowEdit(sheet, editable);

      const fontRange = sheet.getRange(`A${config.dataRowStart}:I${config.dataRowEnd}`);
      const font = fontRange.getFormat().getFont();
      font.setName("Times New Roman");
      font.setBold(true);
      font.setSize(14);
      fontRange.getFormat().setShrinkToFit(true);

      const uppercaseColumns = [
        "PHONE 6.100 AC-PC", "PHONE 6.200 AC-PC",
        "PHONE 7.100 AC-PC", "PHONE 7.200 UNI"
      ].includes(name) ? ["B", "C", "D", "E", "G"] : ["C", "D", "G"];
      makeSelectedColumnsUppercase(sheet, config.dataRowStart, config.dataRowEnd, uppercaseColumns);
    }

    if (state.paused) {
      resumeIfPaused(sheet);
    } else if (wasProtected) {
      try { protection.protect(getProtectionOptions(), password); }
      catch { console.log(`Failed to re-protect "${name}".`); }
    }
  }
}

function createValidationNamedRanges(workbook: ExcelScript.Workbook) {
  const definitions = [
    { name: "RN_NCT_NAMES_E", formula: "='RN-NCT LIST'!$E$2:$E$2000" },
    { name: "RN_NCT_NAMES_F", formula: "='RN-NCT LIST'!$F$2:$F$2000" }
  ];
  for (const definition of definitions) {
    try {
      const existing: ExcelScript.NamedItem = workbook.getNamedItem(definition.name);
      existing.delete();
    } catch {
      // The workbook-level name does not exist yet.
    }

    workbook.addNamedItem(
      definition.name,
      definition.formula,
      "Phone-sheet validation source"
    );
  }
}

function applyFormulas(
  sheet: ExcelScript.Worksheet,
  startRow: number,
  endRow: number
) {
  const is7200UNI: boolean = sheet.getName() === "PHONE 7.200 UNI";

  /*
   * PHONE 7.200 UNI:
   *   Input name in column E
   *   Result written to column F
   *   RN-NCT LIST lookup columns E and F
   *
   * Other non-ICU sheets:
   *   Existing behavior is preserved.
   */
  const names: string = is7200UNI ? "E:E" : "F:F";
  const numbers: string = is7200UNI ? "F:F" : "G:G";

  const b: string = `B${startRow}:B${endRow}`;
  const c: string = `C${startRow}:C${endRow}`;
  const d: string = `D${startRow}:D${endRow}`;

  const countExpression: string =
    `(LEN(${b})-LEN(SUBSTITUTE(${b},"/",""))+1)` +
    `-(LEN(${b})-LEN(SUBSTITUTE(${b},"//","/")))` +
    `-IF(LEFT(${b},1)="/",1,0)` +
    `-IF(RIGHT(${b},1)="/",1,0)` +
    `-((LEN(UPPER(${b}))-LEN(SUBSTITUTE(UPPER(${b}),"*(ADMIT)*","")))/7)` +
    `-((LEN(UPPER(${b}))-LEN(SUBSTITUTE(UPPER(${b}),"*BLOCKED*","")))/9)` +
    `-((LEN(UPPER(${b}))-LEN(SUBSTITUTE(UPPER(${b}),"*RESERVED*","")))/10)`;

  const ac: string =
    `=SUMPRODUCT(` +
    `ISNUMBER(SEARCH("AC",${d}))*` +
    `IF(${b}<>"",(${countExpression}),0)` +
    `)`;

  const pc: string =
    `=SUMPRODUCT(` +
    `ISNUMBER(SEARCH("PC",${d}))*` +
    `IF(${b}<>"",(${countExpression}),0)` +
    `)`;

  const total: string =
    `=SUMPRODUCT(` +
    `IF(${b}<>"",` +
    `(LEN(${b})-LEN(SUBSTITUTE(${b},"/",""))+1)` +
    `-((LEN(UPPER(${b}))-LEN(SUBSTITUTE(UPPER(${b}),"(ADMIT)","")))/7),` +
    `0)` +
    `)`;

  const statusCount = (code: string): string => {
    return (
      `=COUNTIFS(` +
      `${c},"*${code}*",` +
      `${b},"<>*(ADMIT)*",` +
      `${b},"<>*BLOCKED*",` +
      `${b},"<>*RESERVED*"` +
      `)`
    );
  };

  const f: string[][] = [];
  const h: string[][] = [];

  for (let row: number = startRow; row <= endRow; row++) {
    if (is7200UNI) {
      /*
       * PHONE 7.200 UNI column F:
       * Same simple lookup style as ICU, but uses E instead of G.
       */
      f.push([
        `=IF(E${row}<>"",` +
        `IFERROR(` +
        `XLOOKUP(E${row},'RN-NCT LIST'!E:E,'RN-NCT LIST'!F:F),` +
        `""` +
        `),` +
        `""` +
        `)`
      ]);

      /*
       * PHONE 7.200 UNI column H:
       * Reads the name from G and searches RN-NCT LIST E:F.
       */
      h.push([
        `=IF(G${row}<>"",` +
        `IFERROR(` +
        `XLOOKUP(G${row},'RN-NCT LIST'!E:E,'RN-NCT LIST'!F:F),` +
        `""` +
        `),` +
        `""` +
        `)`
      ]);
    } else {
      /*
       * Preserve the existing split-name lookup behavior
       * for all other non-ICU phone sheets.
       */
      f.push([
        `=IF(E${row}="","",` +
        `IF(ISNUMBER(SEARCH("/",E${row})),` +
        `IFERROR(` +
        `XLOOKUP(` +
        `TRIM(LEFT(E${row},FIND("/",E${row})-1)),` +
        `'RN-NCT LIST'!${names},` +
        `'RN-NCT LIST'!${numbers}` +
        `),` +
        `""` +
        `)` +
        `&" / "&` +
        `IFERROR(` +
        `XLOOKUP(` +
        `TRIM(MID(E${row},FIND("/",E${row})+1,999)),` +
        `'RN-NCT LIST'!${names},` +
        `'RN-NCT LIST'!${numbers}` +
        `),` +
        `""` +
        `),` +
        `IFERROR(` +
        `XLOOKUP(` +
        `E${row},` +
        `'RN-NCT LIST'!${names},` +
        `'RN-NCT LIST'!${numbers}` +
        `),` +
        `""` +
        `)` +
        `)` +
        `)`
      ]);

      h.push([
        `=IF(G${row}<>"",` +
        `IFERROR(` +
        `XLOOKUP(` +
        `G${row},` +
        `'RN-NCT LIST'!${names},` +
        `'RN-NCT LIST'!${numbers}` +
        `),` +
        `""` +
        `),` +
        `""` +
        `)`
      ]);
    }
  }

  sheet
    .getRange(`F${startRow}:F${endRow}`)
    .setFormulas(f);

  sheet
    .getRange(`H${startRow}:H${endRow}`)
    .setFormulas(h);

  sheet.getRange("O11").setFormula(ac);
  sheet.getRange("O12").setFormula(pc);
  sheet.getRange("O13").setFormula(total);
  sheet.getRange("O21").setFormula(total);

  /*
   * Charge RN lookup.
   * PHONE 7.200 UNI uses RN-NCT LIST columns E and F.
   * Other sheets retain their existing source columns.
   */
  sheet.getRange("F4").setFormula(
    `=IF(C4<>"",` +
    `IFERROR(` +
    `XLOOKUP(C4,'RN-NCT LIST'!${names},'RN-NCT LIST'!${numbers}),` +
    `""` +
    `),` +
    `""` +
    `)`
  );

  sheet.getRange("O16").setFormula(statusCount("SGE"));
  sheet.getRange("O17").setFormula(statusCount("SGT"));
  sheet.getRange("O18").setFormula(statusCount("SCR"));
  sheet.getRange("O19").setFormula(statusCount("SGO"));

  sheet.getRange("O24").setFormula(
    `=COUNTIFS(` +
    `${b},"*(+)*",` +
    `${b},"<>*(ADMIT)*",` +
    `${b},"<>*BLOCKED*",` +
    `${b},"<>*RESERVED*"` +
    `)`
  );

  sheet.getRange("O27").setFormula(
    `=SUM(` +
    `COUNTIFS(` +
    `${b},{"*(SIT)*","*(TS)*","*(72)*"},` +
    `${b},"<>*(ADMIT)*",` +
    `${b},"<>*BLOCKED*",` +
    `${b},"<>*RESERVED*"` +
    `)` +
    `)`
  );
}

function applyFormulasICU(sheet: ExcelScript.Worksheet, startRow: number, endRow: number) {
  const special = sheet.getName() === "PHONE 7.200 ICU" || sheet.getName() === "PHONE 6.200 ICU";
  const chargeCell = special ? "M2" : "F4";
  const inputCell = special ? "L2" : "C4";
  sheet.getRange(chargeCell).setFormula(`=IF(${inputCell}<>"",IFERROR(XLOOKUP(${inputCell},'RN-NCT LIST'!E:E,'RN-NCT LIST'!F:F),""),"")`);
  const formulas: string[][] = [];
  for (let row = startRow; row <= endRow; row++) formulas.push([`=IF(G${row}<>"",IFERROR(XLOOKUP(G${row},'RN-NCT LIST'!E:E,'RN-NCT LIST'!F:F),""),"")`]);
  sheet.getRange(`I${startRow}:I${endRow}`).setFormulas(formulas);
  for (const cell of ["I10", "I12", "I14", "I16", "I18", "I20"]) sheet.getRange(cell).getFormat().getFill().setColor("FFFFFF");
}

function canonicalName(value: string): string {
  return (value || "").toUpperCase().replace(/\u00A0/g, " ").replace(/\s+/g, " ").replace(/\s*,\s*/g, ",").replace(/\./g, "").trim();
}
function splitIntoNames(value: string): string[] {
  return (value || "").replace(/\/\//g, "/").replace(/\s*\/\s*/g, "/").split("/").map((part: string): string => canonicalName(part)).filter((part: string): boolean => part.length > 0 && !part.includes("RESOURCE"));
}
function writeRNCount(sheet: ExcelScript.Worksheet, startRow: number, endRow: number) {
  const column = sheet.getName().toUpperCase().includes("ICU") ? "G" : "E";
  const values = sheet.getRange(`${column}${startRow}:${column}${endRow}`).getTexts();
  const names = new Set<string>();
  for (const row of values) for (const name of splitIntoNames(row[0])) names.add(name);
  for (const name of splitIntoNames(sheet.getRange("C4").getText())) names.add(name);
  sheet.getRange("O7").setValue(names.size);
}
function writeNCTCount(sheet: ExcelScript.Worksheet, startRow: number, endRow: number) {
  if (sheet.getName().toUpperCase().includes("ICU")) {
    sheet.getRange("O8").clear(ExcelScript.ClearApplyTo.contents);
    return;
  }
  const names = new Set<string>();
  for (const row of sheet.getRange(`G${startRow}:G${endRow}`).getTexts()) for (const name of splitIntoNames(row[0])) names.add(name);
  sheet.getRange("O8").setValue(names.size);
}

function makeSelectedColumnsUppercase(sheet: ExcelScript.Worksheet, startRow: number, endRow: number, columns: string[]) {
  for (const column of columns) {
    const range = sheet.getRange(`${column}${startRow}:${column}${endRow}`);
    const values = range.getValues();
    const formulas = range.getFormulas();
    for (let r = 0; r < values.length; r++) if (formulas[r][0] === "" && typeof values[r][0] === "string") values[r][0] = (values[r][0] as string).toUpperCase();
    range.setValues(values);
  }
}

function applyBorders(sheet: ExcelScript.Worksheet, address: string) {
  const indexes = [ExcelScript.BorderIndex.edgeTop, ExcelScript.BorderIndex.edgeRight, ExcelScript.BorderIndex.edgeLeft, ExcelScript.BorderIndex.edgeBottom, ExcelScript.BorderIndex.insideHorizontal, ExcelScript.BorderIndex.insideVertical];
  const format = sheet.getRange(address).getFormat();
  for (let i = 0; i < indexes.length; i++) {
    const border = format.getRangeBorder(indexes[i]);
    border.setStyle(ExcelScript.BorderLineStyle.continuous);
    border.setWeight(i < 4 ? ExcelScript.BorderWeight.thick : ExcelScript.BorderWeight.thin);
    border.setColor("000000");
  }
}

function condFormatting(sheet: ExcelScript.Worksheet, startRow: number, endRow: number) {
  const range = sheet.getRange(`B${startRow}:B${endRow}`);
  range.getFormat().getFill().clear();
  range.getFormat().getFont().setBold(false);
  range.clearAllConditionalFormats();
  for (const item of [["(S)", "red"], ["(P)", "red"], ["(SPA)", "yellow"], ["(ALS)", "yellow"], ["(INTR)", "yellow"], ["(INT)", "yellow"]]) applyContainsTextCF(range, item[0], item[1]);
  addCustomCF(range, `=AND(INDEX($M:$M,ROW())=2,ISERROR(SEARCH("BLOCKED",INDEX($B:$B,ROW()))),ISERROR(SEARCH("RESERVED",INDEX($B:$B,ROW()))))`, "33ccff", "000000");
  addCustomCF(range, `=OR(ISNUMBER(SEARCH("BLOCKED",INDEX($B:$B,ROW()))),ISNUMBER(SEARCH("RESERVED",INDEX($B:$B,ROW()))))`, "000000", "FFFFFF");
}
function condFormattingICU(sheet: ExcelScript.Worksheet, startRow: number, endRow: number) {
  const range = sheet.getRange(`D${startRow}:D${endRow}`);
  range.clearAllConditionalFormats();
  applyContainsTextCF(range, "(S)", "red");
  applyContainsTextCF(range, "(P)", "red");
  addCustomCF(range, `=AND(INDEX($N:$N,ROW())>1,INDEX($N:$N,ROW())<5,ISERROR(SEARCH("BLOCKED",INDEX($D:$D,ROW()))),ISERROR(SEARCH("RESERVED",INDEX($D:$D,ROW()))))`, "33ccff", "000000");
  addCustomCF(range, `=OR(ISNUMBER(SEARCH("BLOCKED",INDEX($D:$D,ROW()))),ISNUMBER(SEARCH("RESERVED",INDEX($D:$D,ROW()))))`, "000000", "FFFFFF");
}
function applyContainsTextCF(range: ExcelScript.Range, text: string, fill: string) {
  const comparison = range.addConditionalFormat(ExcelScript.ConditionalFormatType.containsText).getTextComparison();
  comparison.setRule({ text, operator: ExcelScript.ConditionalTextOperator.contains });
  comparison.getFormat().getFill().setColor(fill);
  comparison.getFormat().getFont().setBold(true);
}
function addCustomCF(range: ExcelScript.Range, formula: string, fill: string, fontColor: string) {
  const custom = range.addConditionalFormat(ExcelScript.ConditionalFormatType.custom).getCustom();
  custom.getRule().setFormula(formula);
  custom.getFormat().getFill().setColor(fill);
  custom.getFormat().getFont().setColor(fontColor);
  custom.getFormat().getFont().setBold(true);
}

function setListValidation(range: ExcelScript.Range, source: string, title: string, message: string) {
  const validation = range.getDataValidation();
  validation.clear();
  validation.setRule({ list: { source, inCellDropDown: true } });
  validation.setIgnoreBlanks(true);
  validation.setErrorAlert({ title, message, showAlert: true, style: ExcelScript.DataValidationAlertStyle.warning });
}
function dataValidation(sheet: ExcelScript.Worksheet, startRow: number, endRow: number) {
  const source = sheet.getName() === "PHONE 7.200 UNI" ? "=RN_NCT_NAMES_E" : "=RN_NCT_NAMES_F";
  setListValidation(sheet.getRange(`E${startRow}:E${endRow}`), source, "RN not found", "Select an RN from RN-NCT LIST.");
  setListValidation(sheet.getRange(`G${startRow}:G${endRow}`), source, "NCT not found", "Select an NCT from RN-NCT LIST.");
  setListValidation(sheet.getRange(`D${startRow}:D${endRow}`), "AC,PC", "Location not found", "Select AC or PC.");
  setListValidation(sheet.getRange(`C${startRow}:C${endRow}`), "=$B$78:$B$128", "Team not found", "Select a valid team.");
  setListValidation(sheet.getRange("C5"), "7 AM,7 PM", "Invalid shift", "Select 7 AM or 7 PM.");
}
function dataValidationICU(sheet: ExcelScript.Worksheet, startRow: number, endRow: number) {
  setListValidation(sheet.getRange(`G${startRow}:G${endRow}`), "=RN_NCT_NAMES_E", "RN not found", "Select an RN from RN-NCT LIST.");
  setListValidation(sheet.getRange(`C${startRow}:C${endRow}`), "=$B$78:$B$128", "Team not found", "Select a valid team.");
  setListValidation(sheet.getRange("D3"), "7 AM,7 PM", "Invalid shift", "Select 7 AM or 7 PM.");
}

function ensureEditableForWrites(workbook: ExcelScript.Workbook, sheet: ExcelScript.Worksheet, primaryPassword: string): { ok: boolean; paused: boolean } {
  const protection = sheet.getProtection();
  if (!protection.getProtected()) return { ok: true, paused: false };
  const passwords = [primaryPassword, "imsorrydave", "72icu", "61"].filter((value: string, index: number, array: string[]): boolean => array.indexOf(value) === index);
  const passwordProtected = protection.getIsPasswordProtected();
  if (protection.getCanPauseProtection()) {
    if (passwordProtected) {
      for (const password of passwords) {
        try { protection.pauseProtection(password); } catch { }
        if (protection.getIsPaused()) return { ok: true, paused: true };
      }
    } else {
      try { protection.pauseProtection(); } catch { }
      if (protection.getIsPaused()) return { ok: true, paused: true };
    }
  }
  if (passwordProtected) {
    for (const password of passwords) {
      try { protection.unprotect(password); } catch { }
      if (!protection.getProtected()) return { ok: true, paused: false };
    }
  } else {
    try { protection.unprotect(); } catch { }
    if (!protection.getProtected()) return { ok: true, paused: false };
  }
  return { ok: false, paused: false };
}
function resumeIfPaused(sheet: ExcelScript.Worksheet) {
  const protection = sheet.getProtection();
  try { if (protection.getIsPaused()) protection.resumeProtection(); } catch { }
}
function clearAndSetAllowEdit(
  sheet: ExcelScript.Worksheet,
  ranges: string[]
) {
  const protection = sheet.getProtection();

  // Lock whole sheet first
  sheet.getUsedRange()
    ?.getFormat()
    .getProtection()
    .setLocked(true);

  // Unlock editable ranges
  for (const address of ranges) {
    sheet
      .getRange(address)
      .getFormat()
      .getProtection()
      .setLocked(false);
  }
}
function getProtectionOptions(): ExcelScript.WorksheetProtectionOptions {
  return { selectionMode: ExcelScript.ProtectionSelectionMode.unlocked, allowAutoFilter: true, allowSort: true, allowPivotTables: false };
}
function getProtectionPassword(name: string): string {
  if (name === "PHONE 7.200 ICU") return "72icu";
  if (name === "PHONE 6.100 UNI" || name === "PHONE 6.100 ONC") return "61";
  return "imsorrydave";
}

function nameAlerts(sheet: ExcelScript.Worksheet, startRow: number, endRow: number) {
  const values = sheet.getRange(`B${startRow}:B${endRow}`).getTexts();
  const output: number[][] = [];
  const counts: { [key: string]: number } = {};
  const normalizedRows: string[][] = [];
  const clean = (value: string) => value.replace(/\((S|P|SIT|ADMIT)\)/gi, "").trim().toLowerCase().replace(/[^a-z/ ,]/g, "");
  const normalize = (value: string) => value.split(",")[0].trim().replace(/(.)\1+/g, "$1");
  for (const row of values) {
    const raw = row[0];
    if (!raw || /blocked|reserved/i.test(raw)) { normalizedRows.push([]); continue; }
    const names = clean(raw).split("/").map((value: string): string => normalize(value)).filter((value: string): boolean => value !== "");
    normalizedRows.push(names);
    for (const name of names) counts[name] = (counts[name] || 0) + 1;
  }
  for (const names of normalizedRows) output.push([names.length === 0 ? 0 : names.some((name: string): boolean => counts[name] > 1) ? 2 : 1]);
  sheet.getRange(`M${startRow}:M${endRow}`).setValues(output);
}

type SheetConfig = { editableRanges: string[]; dataRowStart: number; dataRowEnd: number };
function getSheetConfig(name: string): SheetConfig {
  const configs: { [key: string]: SheetConfig } = {
    "PHONE 6.100 AC-PC": { editableRanges: ["C4", "C5", "B7:B38", "C7:C38", "D7:D38", "E7:E38", "G7:G38", "L3:L5"], dataRowStart: 7, dataRowEnd: 38 },
    "PHONE 10.200 AC-PC": { editableRanges: ["C4", "C5", "B7:B38", "C7:C38", "D7:D38", "E7:E38", "G7:G38", "L3:L5"], dataRowStart: 7, dataRowEnd: 38 },
    "PHONE 6.200 AC-PC": { editableRanges: ["C4", "C5", "B7:B26", "C7:C26", "D7:D26", "E7:E26", "G7:G26", "L3:L5"], dataRowStart: 7, dataRowEnd: 26 },
    "PHONE 7.200 UNI": { editableRanges: ["C4", "C5", "B7:B26", "C7:C26", "D7:D26", "E7:E26", "G7:G26", "L3:L5"], dataRowStart: 7, dataRowEnd: 26 },
    "PHONE 7.200 ICU": { editableRanges: ["C4", "C5", "C9:C20", "D9:D20", "G9:G20", "Q5:Q7", "J9:J20", "M9:M20"], dataRowStart: 9, dataRowEnd: 20 },
    "PHONE 6.100 UNI": { editableRanges: ["C4", "C5", "B7:B26", "C7:C26", "D7:D26", "E7:E26", "G7:G26", "L3:L5"], dataRowStart: 7, dataRowEnd: 26 },
    "PHONE 6.100 ONC": { editableRanges: ["C4", "C5", "B7:B18", "C7:C18", "D7:D18", "E7:E18", "G7:G18", "L3:L5"], dataRowStart: 7, dataRowEnd: 18 },
    "PHONE 6.200 ICU": { editableRanges: ["L2", "M2", "C9:C20", "D3", "G9:G20", "I9:I20", "Q5:Q7"], dataRowStart: 9, dataRowEnd: 20 }
  };
  const config = configs[name];
  if (!config) throw new Error(`Unsupported sheet: ${name}`);
  return config;
}

function refreshExistingValidation(
  workbook: ExcelScript.Workbook
) {

  const ranges = [
    ["PHONE 6.100 AC-PC", "E7:E38"],
    ["PHONE 6.100 AC-PC", "G7:G38"],

    ["PHONE 6.200 AC-PC", "E7:E26"],
    ["PHONE 6.200 AC-PC", "G7:G26"],

    ["PHONE 10.200 AC-PC", "E7:E38"],
    ["PHONE 10.200 AC-PC", "G7:G38"],

    ["PHONE 6.100 UNI", "E7:E26"],
    ["PHONE 6.100 UNI", "G7:G26"],

    ["PHONE 6.100 ONC", "E7:E18"],
    ["PHONE 6.100 ONC", "G7:G18"],

    ["PHONE 7.200 UNI", "E7:E26"],
    ["PHONE 7.200 UNI", "G7:G26"],

    ["PHONE 6.200 ICU", "G9:G20"],
    ["PHONE 7.200 ICU", "G9:G20"]
  ];

  for (const item of ranges) {

    const sheet =
      workbook.getWorksheet(item[0]);

    if (!sheet) {
      continue;
    }

    const validation =
      sheet
        .getRange(item[1])
        .getDataValidation();

    try {

      const rule =
        validation.getRule();

      if (rule) {
        validation.setRule(rule);
      }

    } catch {
      // ignore
    }
  }
}
