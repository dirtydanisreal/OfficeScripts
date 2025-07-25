
function main(workbook: ExcelScript.Workbook) {
    // Get the active cell and worksheet.
  const sheetNames = ["PHONE 6.100 AC-PC", "PHONE 6.200 AC-PC", "PHONE 7.200 ICU"];

  for (const name of sheetNames) {
    const sheet = workbook.getWorksheet(name);
    if (!sheet) continue;

    const config = getSheetConfig(name);
    const protection = sheet.getProtection();

    protection.pauseProtection("imsorrydave");

      nameAlerts(sheet, config.dataRowStart, config.dataRowEnd);

      protection.resumeProtection();

  }
}

function nameAlerts(sheet: ExcelScript.Worksheet, startRow: number, endRow: number) {

    if(sheet.getName() === "PHONE 7.200 ICU"){
        const range = sheet.getRange(`D${startRow}:D${endRow}`);
    }
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
        if(sheet.getName() === "PHONE 7.200 ICU"){
            const cell = sheet.getRange(`D${startRow + i}`);
        }
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
    } else if(sheetName === "PHONE 7.200 ICU") {
      return {
        range: "A7:I18",
        editableRanges: ["C4", "C5", "B9:B20", "C9:C20", "D9:D10", "E9:E20", "G9:G20", "Q5:Q7"],
        dataRowStart: 9,
        dataRowEnd: 20
    }

    
    } else {
        throw new Error(`Unsupported sheet: ${sheetName}`);
    }
}

