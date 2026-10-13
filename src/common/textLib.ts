export function piece(str: string, delimiter: string, index: number): string {
    const parts = str.split(delimiter);
    if (index <= 0 || index > parts.length) return "";
    return parts[index - 1].trim();
}
export function splitCSV(csv: string): string[] {
    const result: string[] = [];
    let current = "";
    let inQuotes = false;

    csv = csv.trim();

    for (let i = 0; i < csv.length; i++) {
        const char = csv[i];

        if (char === '"') {
            inQuotes = !inQuotes;
            current += char;
        } else if (char === "," && !inQuotes) {
            result.push(current);
            current = "";
        } else {
            current += char;
        }
    }

    if (current !== "") {
        result.push(current);
    }

    return result;
}

// ---------------------------------------------------------
// Utility: Convert ArrayBuffer → Base64
// ---------------------------------------------------------
function arrayBufferToBase64(buffer: ArrayBuffer): string {
    let binary = "";
    const bytes = new Uint8Array(buffer);
    const len = bytes.byteLength;
  
    for (let i = 0; i < len; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
  
    return btoa(binary);
  }
