/**
 * Sets the value of a named range.
 * Creates its own Excel.run context to avoid circular reference issues.
 * @param name Name of the named range
 * @param value Value to set
 * @returns Promise<void>
 */
export async function setNamedRangeValue(name: string, value: any): Promise<void> {
    return Excel.run(async (context) => {
        const namedItem = context.workbook.names.getItem(name);
        const range = namedItem.getRange();
        range.values = [[value]];
        await context.sync();
    });
}

/**
 * Gets the value of a named range.
 * Creates its own Excel.run context to avoid circular reference issues.
 * @param name Name of the named range
 * @returns Promise<any> Value of the named range
 */
export async function getNamedRangeValue(name: string): Promise<any> {
    return Excel.run(async (context) => {
        let namedItem;
        try {
            namedItem = context.workbook.names.getItem(name);
            // Try to access the item to ensure it exists
            namedItem.load("name");
            await context.sync();
        } catch (e) {
            console.warn(`Named range '${name}' does not exist.`);
            return undefined;
        }

        const range = namedItem.getRange();
        range.load("values");
        await context.sync();
        return range.values[0][0];
    });
}