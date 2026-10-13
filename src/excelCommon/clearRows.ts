export async function clearRows(
    sheet: Excel.Worksheet,
    startRow: number = 1,
    clearFormats: boolean = false
) {
    const used = sheet.getUsedRangeOrNullObject();
    used.load(["rowCount", "columnCount"]);
    await sheet.context.sync();

    // If sheet is empty or has no used range, nothing to clear
    if (used.isNullObject) {
        console.log("Sheet is empty or has no used range, nothing to clear");
        return;
    }

    // Convert Excel row number → Office.js row index
    const startIndex = startRow - 1;
    const rowCount = used.rowCount;

    if (rowCount <= startIndex || rowCount <= 0) {
        return;
    }

    const rowsToClear = rowCount - startIndex;
    if (rowsToClear <= 0) {
        return;
    }

    const range = sheet.getRangeByIndexes(
        startIndex,
        0,
        rowsToClear,
        used.columnCount
    );

    // MUST load before reading
    // range.load("address");
    // await sheet.context.sync();

    // console.log("Clearing range:", range.address);

    if (clearFormats) {
        range.clear("All");
    } else {
        range.clear("Contents");
    }
    
    // Sync to ensure the clear operation is executed
    await sheet.context.sync();
}
