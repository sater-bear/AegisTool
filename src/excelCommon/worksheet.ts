export async function goToWorksheet(sheetName: string) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        sheet.activate();
        await context.sync();
    });
}

export async function setFontForWorksheet(sheetName: string, font: string, size: number) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const cells = sheet.getUsedRange();
        cells.format.font.name = font;
        cells.format.font.size = size;
        await context.sync();
    });
}
