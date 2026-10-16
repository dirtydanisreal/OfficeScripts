
// Interfaces
interface LabelRange {
  labelCell: ExcelScript.Range;
  ranges: ExcelScript.Range[];
}

interface TrimmedLabel {
  label: string;                // normalized floor key: "6", "7", "8", "9", "10"
  ranges: ExcelScript.Range[];
}

interface MatchPlan {
  row: number;
  col: number;
  floor: string;
  code: string;
  // prepared outputs
  belowValue: string;  // value to put at (row+1, col)
  col5Value: string;   // value to put at (row, col+2)
  col4Value: string;   // value to put at (row, col+3)
  covidFlag: boolean;
  covidBlack: boolean;
  color: string;       // base block color
}

// --- Helpers ---
function isNumericCellValue(val: string | number | boolean): boolean {
  if (typeof val === "number") return Number.isFinite(val);
  if (typeof val === "boolean") return false;
  const s = String(val ?? "").trim();
  return /^[0-9]+$/.test(s); // digits-only strings treated as room numbers
}

// Function to get label map
function getLabelMap(sheet: ExcelScript.Worksheet): LabelRange[] {
  return [
    { labelCell: sheet.getRange("E14"), ranges: [sheet.getRange("A1:A19"), sheet.getRange("E1:M1"), sheet.getRange("M4:M7")] },
    { labelCell: sheet.getRange("E31"), ranges: [sheet.getRange("A22:A46"), sheet.getRange("E46:I46"), sheet.getRange("M22:M46")] },
    { labelCell: sheet.getRange("E63"), ranges: [sheet.getRange("A50:A95"), sheet.getRange("E95:I95"), sheet.getRange("M50:M56"), sheet.getRange("M59:M66"), sheet.getRange("M71:M95"), sheet.getRange("M68:M68"), sheet.getRange("E50:L50")] },
    { labelCell: sheet.getRange("V14"), ranges: [sheet.getRange("R1:R7"), sheet.getRange("V1:Z1"), sheet.getRange("AD1:AD19")] },
    { labelCell: sheet.getRange("V31"), ranges: [sheet.getRange("R22:R46"), sheet.getRange("V46:Z46"), sheet.getRange("AD22:AD46")] },
    { labelCell: sheet.getRange("V63"), ranges: [sheet.getRange("R50:R56"), sheet.getRange("V50:Z50"), sheet.getRange("M59:M66"), sheet.getRange("AD68:AD68"), sheet.getRange("AD50:AD68")] },
    { labelCell: sheet.getRange("V80"), ranges: [sheet.getRange("R71:R95"), sheet.getRange("V95:Z95"), sheet.getRange("AD71:AD95"), sheet.getRange("T56:U56"), sheet.getRange("R57:R57"), sheet.getRange("T7:U7"), sheet.getRange("R8:R8"), sheet.getRange("C19:D19"), sheet.getRange("A20:A20")] },
    { labelCell: sheet.getRange("AO14"), ranges: [sheet.getRange("AK1:AK7"), sheet.getRange("AO1:AS1"), sheet.getRange("AW1:AW46"), sheet.getRange("AO46:AS46"), sheet.getRange("AK22:AK46"), sheet.getRange("O59:P59"), sheet.getRange("M60:M60"), sheet.getRange("O62:P62"), sheet.getRange("M63:M63"), sheet.getRange("O65:P65"), sheet.getRange("M66:M66"), sheet.getRange("O68:P68"), sheet.getRange("M69:M69"), sheet.getRange("O22:P22"), sheet.getRange("M23:M23"), sheet.getRange("O25:P25"), sheet.getRange("M26:M26"), sheet.getRange("O31:P31"), sheet.getRange("M32:M32")] }
  ];
}

/**
 * Extract a normalized floor key from text:
 * - pulls the first numeric substring
 * - only accepts floors 6, 7, 8, 9, 10
 * - returns "" if the text doesn't represent a floor in [6..10]
 */
function extractFloorKey(text: string): string {
  const clean = (text ?? "").trim();
  const m = clean.match(/\d+/);
  if (!m) return "";
  const num = parseInt(m[0], 10);
  if (num >= 6 && num <= 10) return String(num); // allow two-digit "10" and single-digit 6..9
  return "";
}

// Function to trim labels (normalize to floor keys 6..10)
function trimLabels(labelMap: LabelRange[]): TrimmedLabel[] {
  return labelMap.map((labelRange: LabelRange): TrimmedLabel => {
    const labelText = String(labelRange.labelCell.getText() ?? "");
    const trimmed = extractFloorKey(labelText);
    return {
      label: trimmed,
      ranges: labelRange.ranges
    };
  });
}

/** Build a lookup map from CHARGE_SHEET: floor_code -> {row,col} */
function buildCodeMap(trimmedLabelMap: TrimmedLabel[]): Map<string, { row: number; col: number }> {
  const codeMap: Map<string, { row: number; col: number }> = new Map();
  const duplicateKeys: string[] = [];

  trimmedLabelMap.forEach((trimmedLabel: TrimmedLabel) => {
    if (!trimmedLabel.label) return;

    trimmedLabel.ranges.forEach((range: ExcelScript.Range) => {
      let values: (string | number | boolean)[][];
      try {
        values = range.getValues() as (string | number | boolean)[][];
      } catch {
        return; // guard: skip problematic range
      }
      const startRow = range.getRowIndex();
      const startCol = range.getColumnIndex();

      values.forEach((row: (string | number | boolean)[], r: number) => {
        row.forEach((cellVal: string | number | boolean, c: number) => {
          const code = String(cellVal).trim();
          if (code) {
            const key = `${trimmedLabel.label}_${code}`;
            if (codeMap.has(key)) duplicateKeys.push(key);
            codeMap.set(key, { row: startRow + r, col: startCol + c });
          }
        });
      });
    });
  });

  // (Optional) you can surface duplicateKeys to LIST_DATA if you want
  return codeMap;
}

/** Normalize unit & column 5 */
function normalizeUnit(unitRaw: string): string {
  if (unitRaw === "Progressive") return "PC";
  if (unitRaw === "Acute") return "AC";
  return unitRaw || "ICU";
}
function normalizeCol5(s: string): string {
  if (s.length > 3 && !["HM", "SGT", "SCR", "SGE"].includes(s)) {
    return s.substring(0, 3);
  }
  return s;
}

/** Build all match plans from LIST_DATA once (fast) */
function buildMatchPlans(
  listRangeValues: (string | number | boolean)[][],
  codeMap: Map<string, { row: number; col: number }>
): { plans: MatchPlan[]; unmatchedLogs: { rowIndex: number; message: string }[] } {
  const plans: MatchPlan[] = [];
  const unmatchedLogs: { rowIndex: number; message: string }[] = [];

  listRangeValues.forEach((listRow: (string | number | boolean)[], index: number) => {
    let [
      listLabel,
      listCode,
      listValueToCopy,
      listFourthColumnValue,
      listFifthColumnValue,
      listSixthColumnValue,
      listSeventhColumnValue,
      listEighthColumnValue
    ] = listRow.map((cell) => String(cell).trim());

    const floor = extractFloorKey(listLabel);
    if (!floor) {
      unmatchedLogs.push({ rowIndex: index, message: `Label not in allowed floors (6-10): ${listLabel}` });
      return;
    }

    const key = `${floor}_${listCode}`;
    const loc = codeMap.get(key);
    if (!loc) {
      unmatchedLogs.push({ rowIndex: index, message: `No match for label: ${floor} / code: ${listCode}` });
      return;
    }

    const col4 = normalizeUnit(listFourthColumnValue);
    const col5 = normalizeCol5(listFifthColumnValue);
    const color =
      col4 === "ICU" ? "4d94db" :
      col4 === "PC"  ? "a6c9ec" :
      "ffffff";

    const belowValue = listSeventhColumnValue === "Yes"
      ? `${listValueToCopy} (Sit)`
      : listValueToCopy;

    plans.push({
      row: loc.row,
      col: loc.col,
      floor,
      code: listCode,
      belowValue,
      col5Value: col5,
      col4Value: col4,
      covidFlag: (listEighthColumnValue || "").toLowerCase().includes("covid"),
      covidBlack: listSixthColumnValue === "Yes" && listEighthColumnValue !== "",
      color
    });
  });

  return { plans, unmatchedLogs };
}

/** Optimize: pre-clear neighbor cells for today's matches (dedup once), BUT NEVER clear numeric cells */
function preClearNeighborsForPlans(sheet: ExcelScript.Worksheet, plans: MatchPlan[]) {
  const seen = new Set<string>();
  const neighbors: Array<{ r: number; c: number }> = [];

  const key = (r: number, c: number) => `${r},${c}`;

  for (const p of plans) {
    const baseR = p.row, baseC = p.col;
    const coords: Array<[number, number]> = [
      [baseR + 1, baseC],
      [baseR + 2, baseC],
      [baseR, baseC + 1],
      [baseR, baseC + 2],
      [baseR, baseC + 3]
    ];
    for (const [r, c] of coords) {
      const k = key(r, c);
      if (!seen.has(k)) {
        seen.add(k);
        neighbors.push({ r, c });
      }
    }
  }

  // Clear neighbors safely (skip if current value is numeric/digits-only)
  for (const n of neighbors) {
    try {
      const cell = sheet.getRangeByIndexes(n.r, n.c, 1, 1);
      const current = cell.getValue() as (string | number | boolean);
      if (isNumericCellValue(current)) {
        // Do NOT clear numeric room/code cells
        continue;
      }
      cell.setValue("");
      const fmt = cell.getFormat();
      fmt.getFill().setColor("ffffff");
      const f = fmt.getFont();
      f.setColor("000000");
      f.setBold(false);
    } catch {
      // out-of-bounds or merged — skip
    }
  }
}

/** Perform value updates and coloring for all plans */
function applyPlans(sheet: ExcelScript.Worksheet, plans: MatchPlan[]) {
  // Value updates
  for (const p of plans) {
    try { sheet.getRangeByIndexes(p.row + 1, p.col, 1, 1).setValue(p.belowValue); } catch {}
    try { sheet.getRangeByIndexes(p.row, p.col + 2, 1, 1).setValue(p.col5Value); } catch {}
    try { sheet.getRangeByIndexes(p.row, p.col + 3, 1, 1).setValue(p.col4Value); } catch {}

    // Base code cell: ensure white fill (keep room number value)
    try { sheet.getRangeByIndexes(p.row, p.col, 1, 1).getFormat().getFill().setColor("ffffff"); } catch {}

    // COVID flag
    if (p.covidFlag) {
      try {
        const covidCell = sheet.getRangeByIndexes(p.row, p.col + 1, 1, 1);
        covidCell.setValue("C");
        const fmt = covidCell.getFormat();
        if (p.covidBlack) {
          fmt.getFill().setColor("black");
          fmt.getFont().setColor("white");
        } else {
          fmt.getFill().setColor(p.color);
          fmt.getFont().setColor("black");
        }
        fmt.getFont().setBold(true);
      } catch {}
    }
  }

  // Coloring (group by color)
  const groups: Record<string, Array<[number, number]>> = {};
  for (const p of plans) {
    const color = p.color;
    if (!groups[color]) groups[color] = [];
    const cells: Array<[number, number]> = [
      [p.row, p.col],
      [p.row, p.col + 1],
      [p.row, p.col + 2],
      [p.row, p.col + 3],
      [p.row + 1, p.col],
      [p.row + 2, p.col]
    ];
    groups[color].push(...cells);
  }

  Object.keys(groups).forEach((color: string) => {
    for (const [r, c] of groups[color]) {
      try { sheet.getRangeByIndexes(r, c, 1, 1).getFormat().getFill().setColor(color); } catch {}
    }
  });
}

/** Write unmatched diagnostics to LIST_DATA column I */
function writeUnmatchedDiagnostics(list: ExcelScript.Worksheet, unmatchedLogs: { rowIndex: number; message: string }[]) {
  unmatchedLogs.forEach((item) => {
    try { list.getRangeByIndexes(item.rowIndex, 8, 1, 1).setValue(item.message); } catch {}
  });
}

// Main
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet("CHARGE_SHEET");
  const list = workbook.getWorksheet("LIST_DATA");
  sheet.getProtection().pauseProtection("dcn");

  // 1) Map & normalize floors
  const labelMap = getLabelMap(sheet);
  const trimmedLabelMap = trimLabels(labelMap);

  // 2) Build code map (from CHARGE_SHEET)
  const codeMap = buildCodeMap(trimmedLabelMap);

  // 3) Load LIST_DATA once
  const listUsed = list.getUsedRange();
  const listRangeValues = listUsed
    ? (listUsed.getValues() as (string | number | boolean)[][])
    : [];

  // 4) Build all match plans (fast, single pass)
  const { plans, unmatchedLogs } = buildMatchPlans(listRangeValues, codeMap);

  // 5) Clear neighbors only for today's matched codes (dedup), but NEVER clear numeric cells
  preClearNeighborsForPlans(sheet, plans);

  // 6) Apply value updates, COVID flags, and colors
  applyPlans(sheet, plans);

  // 7) Diagnostics
  writeUnmatchedDiagnostics(list, unmatchedLogs);

  sheet.getProtection().resumeProtection();
}
