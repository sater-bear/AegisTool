export async function autoFitAllColumns(sheetName: string) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        sheet.getUsedRange().format.autofitColumns();
        await context.sync();
    });
}
