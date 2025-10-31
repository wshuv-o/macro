function main(workbook: ExcelScript.Workbook) {
    // Use active worksheet (or replace with workbook.getWorksheet("Sheet1") if you want a specific one)
    const sheet: ExcelScript.Worksheet = workbook.getActiveWorksheet();

    const xHeader: string = "Leased SF";
    const yHeader: string = "12900 W Airport Blvd";
    const xRangeAddress: string = "A1:C300";
    const yRangeAddress: string = "A1:H300";

    const result: string | number | boolean | null =
        getValueAtIntersectionString(sheet, xHeader, yHeader, xRangeAddress, yRangeAddress);

    if (result !== null && result !== undefined) {
        console.log(result);
    } else {
       console.log("Value not found.");
    }
}

function getValueAtIntersectionString(
    sheet: ExcelScript.Worksheet,
    xHeader: string,
    yHeader: string,
    xRangeAddress: string,
    yRangeAddress: string
): string | number | boolean | null {
    const xRange: ExcelScript.Range = sheet.getRange(xRangeAddress);
    const yRange: ExcelScript.Range = sheet.getRange(yRangeAddress);

    const xValues: (string | number | boolean)[][] = xRange.getValues() as (string | number | boolean)[][];
    const yValues: (string | number | boolean)[][] = yRange.getValues() as (string | number | boolean)[][];

    let xCol: number | null = null;

    // Find xHeader position
    for (let i = 0; i < xValues.length; i++) {
        for (let j = 0; j < xValues[i].length; j++) {
            if (xValues[i][j] === xHeader) {
                xCol = j;
                break;
            }
        }
        if (xCol !== null) break;
    }

    if (xCol === null) {
        console.log("X Header Not Found");
        return null;
    }

    let yRow: number | null = null;

    // Find yHeader position
    for (let i = 0; i < yValues.length; i++) {
        for (let j = 0; j < yValues[i].length; j++) {
            if (yValues[i][j] === yHeader) {
                yRow = i;
                break;
            }
        }
        if (yRow !== null) break;
    }

    if (yRow === null) {
        console.log("Y Header Not Found");
        return null;
    }

    // Get intersection cell
    const resultCell: ExcelScript.Range = sheet.getCell(yRow, xCol);
    const value: string | number | boolean | null = resultCell.getValue() as string | number | boolean | null;

    return value;
}
