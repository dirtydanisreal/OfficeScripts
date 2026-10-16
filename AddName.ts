function main(workbook: ExcelScript.Workbook) {
  // --------------------------------------------------
  // Known worksheet-protection passwords
  // --------------------------------------------------

  const KNOWN_PASSWORDS: string[] = [
    "61",
    "72icu",
    "imsorrydave",
    ""
  ];

  // --------------------------------------------------
  // Get the worksheet from which the script was run
  // --------------------------------------------------

  const sheet: ExcelScript.Worksheet =
    workbook.getActiveWorksheet();

  const sourceSheetName: string =
    sheet.getName();

  // --------------------------------------------------
  // Pause protection on the source worksheet
  // --------------------------------------------------

  const sourceProtection:
    ExcelScript.WorksheetProtection =
    sheet.getProtection();

  const sourceProtectionPaused: boolean =
    pauseSheetProtection(
      sourceProtection,
      KNOWN_PASSWORDS
    );

  try {
    // --------------------------------------------------
    // Read the source input cells
    // --------------------------------------------------

    let firstNameRaw: string =
      toStr(
        sheet
          .getRange("L3")
          .getValue()
      );

    let lastNameRaw: string =
      toStr(
        sheet
          .getRange("L4")
          .getValue()
      );

    let phoneNumberRaw: string =
      toStr(
        sheet
          .getRange("L5")
          .getValue()
      );

    // ICU sheets use Q5, Q6, and Q7.
    if (
      sourceSheetName === "PHONE 6.200 ICU" ||
      sourceSheetName === "PHONE 7.200 ICU"
    ) {
      firstNameRaw =
        toStr(
          sheet
            .getRange("Q5")
            .getValue()
        );

      lastNameRaw =
        toStr(
          sheet
            .getRange("Q6")
            .getValue()
        );

      phoneNumberRaw =
        toStr(
          sheet
            .getRange("Q7")
            .getValue()
        );
    }

    // LEAN PHONE SHEET uses N4, N6, and N8.
    if (
      sourceSheetName === "LEAN PHONE SHEET"
    ) {
      firstNameRaw =
        toStr(
          sheet
            .getRange("N4")
            .getValue()
        );

      lastNameRaw =
        toStr(
          sheet
            .getRange("N6")
            .getValue()
        );

      phoneNumberRaw =
        toStr(
          sheet
            .getRange("N8")
            .getValue()
        );
    }

    // --------------------------------------------------
    // Normalize the entered information
    // --------------------------------------------------

    const firstNameKey: string =
      toKey(firstNameRaw);

    const lastNameKey: string =
      toKey(lastNameRaw);

    const phoneNumberDigits: string =
      digits(phoneNumberRaw);

    const phoneNumberKey: string =
      phoneNumberDigits.length > 5
        ? phoneNumberDigits
        : "";

    if (
      !firstNameKey &&
      !lastNameKey &&
      !phoneNumberKey
    ) {
      throw new Error(
        "No values provided. Enter First, Last, " +
        "and/or Number before running."
      );
    }

    // --------------------------------------------------
    // Find RN-NCT LIST
    // --------------------------------------------------

    const destinationSheet:
      ExcelScript.Worksheet | undefined =
      workbook.getWorksheet("RN-NCT LIST");

    if (!destinationSheet) {
      throw new Error(
        'Worksheet "RN-NCT LIST" was not found.'
      );
    }

    const destinationProtection:
      ExcelScript.WorksheetProtection =
      destinationSheet.getProtection();

    const destinationProtectionPaused: boolean =
      pauseSheetProtection(
        destinationProtection,
        KNOWN_PASSWORDS
      );

    try {
      // --------------------------------------------------
      // Locate the table on RN-NCT LIST
      // --------------------------------------------------

      const tables: ExcelScript.Table[] =
        destinationSheet.getTables();

      if (tables.length === 0) {
        throw new Error(
          "RN-NCT LIST must contain an Excel table."
        );
      }

      const table: ExcelScript.Table =
        tables[0];

      const tableColumns:
        ExcelScript.TableColumn[] =
        table.getColumns();

      if (tableColumns.length < 4) {
        throw new Error(
          "RN-NCT LIST table must contain at least " +
          "four columns through column D."
        );
      }

      // --------------------------------------------------
      // Delete existing matching entries
      // --------------------------------------------------

      const existingRowCount: number =
        table.getRowCount();

      if (existingRowCount > 0) {
        const tableBodyRange: ExcelScript.Range =
          table.getRangeBetweenHeaderAndTotal();

        const tableBodyValues:
          (
            string |
            number |
            boolean
          )[][] =
          tableBodyRange.getValues();

        const rowsToDelete: number[] = [];

        for (
          let rowIndex: number = 0;
          rowIndex < tableBodyValues.length;
          rowIndex++
        ) {
          const existingFirstName: string =
            toKey(
              tableBodyValues[rowIndex][0]
            );

          const existingLastName: string =
            toKey(
              tableBodyValues[rowIndex][1]
            );

          const existingPhoneNumber: string =
            digits(
              tableBodyValues[rowIndex][3]
            );

          let isMatchingRow: boolean = false;

          // ICU records match by first and last name.
          if (
            sourceSheetName === "PHONE 6.200 ICU" ||
            sourceSheetName === "PHONE 7.200 ICU"
          ) {
            if (
              firstNameKey !== "" &&
              lastNameKey !== "" &&
              existingFirstName === firstNameKey &&
              existingLastName === lastNameKey
            ) {
              isMatchingRow = true;
            }
          } else {
            /*
             * Other worksheets first match by phone number.
             * If no usable phone number was provided, they
             * match by first and last name.
             */
            if (
              phoneNumberKey !== "" &&
              existingPhoneNumber === phoneNumberKey
            ) {
              isMatchingRow = true;
            } else if (
              firstNameKey !== "" &&
              lastNameKey !== "" &&
              existingFirstName === firstNameKey &&
              existingLastName === lastNameKey
            ) {
              isMatchingRow = true;
            }
          }

          if (isMatchingRow) {
            rowsToDelete.push(rowIndex);
          }
        }

        /*
         * Delete from bottom to top so deleting one row
         * does not change the indexes of the remaining rows.
         */
        for (
          let deleteIndex: number =
            rowsToDelete.length - 1;

          deleteIndex >= 0;
          deleteIndex--
        ) {
          table.deleteRowsAt(
            rowsToDelete[deleteIndex],
            1
          );
        }
      }

      // --------------------------------------------------
      // Add the new table row
      // --------------------------------------------------

      table.addRow(-1);

      const newRowIndex: number =
        Math.max(
          0,
          table.getRowCount() - 1
        );

      const firstNameColumnRange:
        ExcelScript.Range =
        tableColumns[0]
          .getRangeBetweenHeaderAndTotal();

      const lastNameColumnRange:
        ExcelScript.Range =
        tableColumns[1]
          .getRangeBetweenHeaderAndTotal();

      const phoneNumberColumnRange:
        ExcelScript.Range =
        tableColumns[3]
          .getRangeBetweenHeaderAndTotal();

      firstNameColumnRange
        .getCell(newRowIndex, 0)
        .setValue(firstNameKey);

      lastNameColumnRange
        .getCell(newRowIndex, 0)
        .setValue(lastNameKey);

      phoneNumberColumnRange
        .getCell(newRowIndex, 0)
        .setValue(
          formatPhoneForOutput(
            phoneNumberRaw
          )
        );

     
    } finally {
      // --------------------------------------------------
      // Restore RN-NCT LIST protection
      // --------------------------------------------------

      if (destinationProtectionPaused) {
        try {
          destinationProtection
            .resumeProtection();
        } catch {
          console.log(
            'Could not resume protection on ' +
            '"RN-NCT LIST".'
          );
        }
      }
    }

    // --------------------------------------------------
    // Refresh existing data-validation rules
    // --------------------------------------------------

    /*
     * This function does not create data validation.
     *
     * It does not change an existing validation source.
     *
     * It checks each configured RN/NCT dropdown cell.
     * If the cell already has validation, the same rule is
     * reapplied. If the cell has no validation, it is left
     * completely unchanged.
     */
    refreshActiveSheetValidation(
      workbook,
      KNOWN_PASSWORDS
    );


    // --------------------------------------------------
    // Clear source entry cells
    // --------------------------------------------------

    if (
      sourceSheetName === "PHONE 6.200 ICU" ||
      sourceSheetName === "PHONE 7.200 ICU"
    ) {
      sheet
        .getRange("Q5:Q7")
        .clear(
          ExcelScript.ClearApplyTo.contents
        );
    } else if (
      sourceSheetName === "LEAN PHONE SHEET"
    ) {
      sheet
        .getRange("N4")
        .clear(
          ExcelScript.ClearApplyTo.contents
        );

      sheet
        .getRange("N6")
        .clear(
          ExcelScript.ClearApplyTo.contents
        );

      sheet
        .getRange("N8")
        .clear(
          ExcelScript.ClearApplyTo.contents
        );
    } else {
      sheet.getRange("L3").setValue("");
      sheet.getRange("L4").setValue("");
      sheet.getRange("L5").setValue("");
    }
  } finally {
    // --------------------------------------------------
    // Restore source worksheet protection
    // --------------------------------------------------

    if (sourceProtectionPaused) {
      try {
        sourceProtection
          .resumeProtection();
      } catch {
        console.log(
          `Could not resume protection on ` +
          `"${sourceSheetName}".`
        );
      }
    }
  }
}

// ==================================================
// Refresh existing data validation
// ==================================================

function refreshExistingValidation(
  workbook: ExcelScript.Workbook,
  knownPasswords: string[]
) {
  /*
   * These are only the RN and NCT dropdown ranges.
   *
   * The script does not refresh location, shift, team,
   * or any other validation rules.
   */
  const validationTargets:
    ValidationTarget[] = [
      {
        sheetName: "PHONE 6.100 AC-PC",
        rangeAddress: "E7:E38"
      },
      {
        sheetName: "PHONE 6.100 AC-PC",
        rangeAddress: "G7:G38"
      },

      {
        sheetName: "PHONE 6.200 AC-PC",
        rangeAddress: "E7:E26"
      },
      {
        sheetName: "PHONE 6.200 AC-PC",
        rangeAddress: "G7:G26"
      },

      {
        sheetName: "PHONE 10.200 AC-PC",
        rangeAddress: "E7:E38"
      },
      {
        sheetName: "PHONE 10.200 AC-PC",
        rangeAddress: "G7:G38"
      },

      {
        sheetName: "PHONE 6.100 UNI",
        rangeAddress: "E7:E26"
      },
      {
        sheetName: "PHONE 6.100 UNI",
        rangeAddress: "G7:G26"
      },

      {
        sheetName: "PHONE 6.100 ONC",
        rangeAddress: "E7:E18"
      },
      {
        sheetName: "PHONE 6.100 ONC",
        rangeAddress: "G7:G18"
      },

      {
        sheetName: "PHONE 7.200 UNI",
        rangeAddress: "E7:E26"
      },
      {
        sheetName: "PHONE 7.200 UNI",
        rangeAddress: "G7:G26"
      },

      {
        sheetName: "PHONE 6.200 ICU",
        rangeAddress: "G9:G20"
      },

      {
        sheetName: "PHONE 7.200 ICU",
        rangeAddress: "G9:G20"
      }
    ];

  /*
   * Build a unique list of worksheets so each worksheet's
   * protection is paused and resumed only one time.
   */
  const targetSheetNames: string[] = [];

  for (
    let targetIndex: number = 0;
    targetIndex < validationTargets.length;
    targetIndex++
  ) {
    const target: ValidationTarget =
      validationTargets[targetIndex];

    if (
      targetSheetNames.indexOf(
        target.sheetName
      ) === -1
    ) {
      targetSheetNames.push(
        target.sheetName
      );
    }
  }

  // --------------------------------------------------
  // Process each PHONE worksheet
  // --------------------------------------------------

  for (
    let sheetIndex: number = 0;
    sheetIndex < targetSheetNames.length;
    sheetIndex++
  ) {
    const targetSheetName: string =
      targetSheetNames[sheetIndex];

    const targetSheet:
      ExcelScript.Worksheet | undefined =
      workbook.getWorksheet(
        targetSheetName
      );

    if (!targetSheet) {
      console.log(
        `Validation refresh skipped because ` +
        `"${targetSheetName}" was not found.`
      );

      continue;
    }

    const targetProtection:
      ExcelScript.WorksheetProtection =
      targetSheet.getProtection();

    let targetProtectionPaused: boolean =
      false;

    try {
      targetProtectionPaused =
        pauseSheetProtection(
          targetProtection,
          knownPasswords
        );
    } catch (error) {
      console.log(
        `Validation refresh skipped on ` +
        `"${targetSheetName}". ` +
        `${errorToString(error)}`
      );

      continue;
    }

    try {
      for (
        let targetIndex: number = 0;
        targetIndex < validationTargets.length;
        targetIndex++
      ) {
        const target: ValidationTarget =
          validationTargets[targetIndex];

        if (
          target.sheetName !==
          targetSheetName
        ) {
          continue;
        }

        const validationRange:
          ExcelScript.Range =
          targetSheet.getRange(
            target.rangeAddress
          );

        refreshValidationRangeByCell(
          validationRange
        );
      }
    } finally {
      /*
       * If this function paused the worksheet, this
       * function is also responsible for resuming it.
       */
      if (targetProtectionPaused) {
        try {
          targetProtection
            .resumeProtection();
        } catch {
          console.log(
            `Could not resume protection on ` +
            `"${targetSheetName}" after refreshing ` +
            `data validation.`
          );
        }
      }
    }
  }
}

// ==================================================
// Refresh every cell in one configured range
// ==================================================

function refreshValidationRangeByCell(
  validationRange: ExcelScript.Range
) {
  const rowCount: number =
    validationRange.getRowCount();

  const columnCount: number =
    validationRange.getColumnCount();

  /*
   * Checking each cell avoids the need for a nonexistent
   * DataValidationType.mixed enum value.
   */
  for (
    let rowIndex: number = 0;
    rowIndex < rowCount;
    rowIndex++
  ) {
    for (
      let columnIndex: number = 0;
      columnIndex < columnCount;
      columnIndex++
    ) {
      const validationCell:
        ExcelScript.Range =
        validationRange.getCell(
          rowIndex,
          columnIndex
        );

      refreshValidationCell(
        validationCell
      );
    }
  }
}

// ==================================================
// Refresh one cell only if validation already exists
// ==================================================

function refreshValidationCell(
  validationCell: ExcelScript.Range
) {
  try {
    const dataValidation:
      ExcelScript.DataValidation =
      validationCell.getDataValidation();

    const validationType:
      ExcelScript.DataValidationType =
      dataValidation.getType();

    /*
     * The cell has no validation. Leave it unchanged.
     */
    if (
      validationType ===
      ExcelScript.DataValidationType.none
    ) {
      return;
    }

    /*
     * Retrieve and reapply the exact existing rule.
     * This does not substitute a new validation source.
     */
    const existingRule:
      ExcelScript.DataValidationRule =
      dataValidation.getRule();

    dataValidation.setRule(
      existingRule
    );
  } catch {
    /*
     * If Excel cannot inspect or reapply the rule,
     * leave the cell unchanged.
     */
  }
}

// ==================================================
// Worksheet-protection helper
// ==================================================

function pauseSheetProtection(
  protection: ExcelScript.WorksheetProtection,
  passwords: string[]
): boolean {
  /*
   * Worksheet is not protected, so there is nothing
   * for this function to pause.
   */
  if (!protection.getProtected()) {
    return false;
  }

  /*
   * The worksheet may already be paused because it is
   * also the source worksheet. In that case, do not
   * attempt to pause or later resume it from this helper.
   */
  try {
    const alreadyPaused: boolean =
      protection.getIsPaused();

    if (alreadyPaused) {
      return false;
    }
  } catch {
    /*
     * Continue with the normal pause attempt.
     */
  }

  const canPauseProtection: boolean =
    protection.getCanPauseProtection();

  if (!canPauseProtection) {
    throw new Error(
      "Worksheet protection cannot be paused."
    );
  }

  const passwordProtected: boolean =
    protection.getIsPasswordProtected();

  /*
   * Pause protection without a password when the
   * worksheet is not password protected.
   */
  if (!passwordProtected) {
    try {
      protection.pauseProtection();
    } catch (error) {
      throw new Error(
        `Could not pause worksheet protection: ` +
        `${errorToString(error)}`
      );
    }

    const protectionWasPaused: boolean =
      protection.getIsPaused();

    if (!protectionWasPaused) {
      throw new Error(
        "Excel did not pause worksheet protection."
      );
    }

    return true;
  }

  /*
   * Password-protected worksheet.
   * Try each known password.
   */
  for (
    let passwordIndex: number = 0;
    passwordIndex < passwords.length;
    passwordIndex++
  ) {
    const password: string =
      passwords[passwordIndex];

    try {
      const passwordMatches: boolean =
        protection.checkPassword(
          password
        );

      if (!passwordMatches) {
        continue;
      }

      protection.pauseProtection(
        password
      );

      const protectionWasPaused: boolean =
        protection.getIsPaused();

      if (protectionWasPaused) {
        return true;
      }
    } catch {
      /*
       * Continue trying the other known passwords.
       */
    }
  }

  throw new Error(
    "Could not pause worksheet protection because " +
    "none of the supplied passwords matched."
  );
}

// ==================================================
// String helpers
// ==================================================

function toStr(
  value: unknown
): string {
  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  return String(value).trim();
}

function toKey(
  value: unknown
): string {
  return toStr(value).toUpperCase();
}

function digits(
  value: unknown
): string {
  return toStr(value).replace(
    /\D+/g,
    ""
  );
}

// ==================================================
// Phone-number formatting
// ==================================================

function formatPhoneForOutput(
  rawValue: unknown
): string {
  const phoneDigits: string =
    digits(rawValue);

  /*
   * Seven-digit numbers are stored as 000-0000.
   */
  if (phoneDigits.length === 7) {
    return (
      `${phoneDigits.slice(0, 3)}-` +
      `${phoneDigits.slice(3)}`
    );
  }

  /*
   * Preserve the original value for other lengths.
   */
  return toStr(rawValue);
}

// ==================================================
// Error formatting
// ==================================================

function errorToString(
  error: unknown
): string {
  if (
    typeof error === "object" &&
    error !== null &&
    "message" in error
  ) {
    const errorObject:
      { message?: unknown } =
      error as { message?: unknown };

    return toStr(
      errorObject.message
    );
  }

  return toStr(error);
}

// ==================================================
// Script-specific interfaces
// ==================================================

interface ValidationTarget {
  sheetName: string;
  rangeAddress: string;
}

function refreshActiveSheetValidation(
  workbook: ExcelScript.Workbook,
  knownPasswords: string[]
) {

  const sheet =
    workbook.getActiveWorksheet();

  const sheetName =
    sheet.getName();

  let targets: string[] = [];

  switch (sheetName) {

    case "PHONE 6.100 AC-PC":
      targets = ["E7:E38", "G7:G38"];
      break;

    case "PHONE 6.200 AC-PC":
      targets = ["E7:E26", "G7:G26"];
      break;

    case "PHONE 10.200 AC-PC":
      targets = ["E7:E38", "G7:G38"];
      break;

    case "PHONE 6.100 UNI":
      targets = ["E7:E26", "G7:G26"];
      break;

    case "PHONE 6.100 ONC":
      targets = ["E7:E18", "G7:G18"];
  }

}
