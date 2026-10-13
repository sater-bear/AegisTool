/**
 * Safely extracts error message from error objects without triggering circular reference issues.
 * This is particularly important for Excel API errors which contain circular references.
 * @param err The error object (can be any type)
 * @returns A safe string representation of the error message
 */
export function getErrorMessage(err: any): string {
    try {
        if (err && typeof err === "object") {
            if (typeof err.message === "string") {
                return err.message;
            } else if (typeof err.error === "string") {
                return err.error;
            } else if (typeof err.toString === "function") {
                return err.toString();
            }
            return "Office.js error (could not extract message)";
        } else if (err != null) {
            return String(err);
        }
        return "Unknown error";
    } catch {
        return "Error occurred (details unavailable)";
    }
}
