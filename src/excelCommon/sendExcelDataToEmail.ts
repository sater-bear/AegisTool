/**
 * Sends the active worksheet's used range as an HTML table to a Power Automate flow (email).
 * Table styling reflects Excel format: fill colors, font name/size/color/bold.
 */

// --- Constants: configure these for your flow and email ---
const FLOW_URL =
  "https://defaultfcb6fca8f3c7458b9a80211f5ce84b.61.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/d2347d905f904fcab6a7111a25b2d8e0/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=VrMbrCseOphfZrttLPKm-3whsAo8O58_z7NbXoL3VN0";
const PAYLOAD_RECIPIENT = "your-email@company.com";
const PAYLOAD_SUBJECT = "Excel Data Export";

export async function sendExcelDataToEmail(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();

      range.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      const rowCount = range.rowCount;
      const colCount = range.columnCount;
      const values = (range.values as (string | number)[][]) || [];

      if (rowCount === 0 || colCount === 0) {
        console.warn("No data in used range.");
        return;
      }

      // Load format (fill, font) for each cell so the HTML table matches the sheet
      const cellRanges: Excel.Range[] = [];
      for (let i = 0; i < rowCount; i++) {
        for (let j = 0; j < colCount; j++) {
          const cell = range.getCell(i, j);
          cell.format.fill.load("color");
          cell.format.font.load(["name", "size", "bold", "italic", "color"]);
          cellRanges.push(cell);
        }
      }
      await context.sync();

      // Build 2D array of format info (row, col) -> { fillColor, fontName, fontSize, fontBold, fontItalic, fontColor }
      const formats: {
        fillColor: string;
        fontName: string;
        fontSize: number;
        fontBold: boolean;
        fontItalic: boolean;
        fontColor: string;
      }[][] = [];
      let idx = 0;
      for (let i = 0; i < rowCount; i++) {
        formats[i] = [];
        for (let j = 0; j < colCount; j++) {
          const r = cellRanges[idx++];
          const fill = r.format.fill;
          const font = r.format.font;
          formats[i][j] = {
            fillColor: fill.color && fill.color !== "" ? fill.color : "",
            fontName: font.name || "Calibri",
            fontSize: font.size ? Math.round(font.size) : 11,
            fontBold: font.bold || false,
            fontItalic: font.italic || false,
            fontColor: font.color && font.color !== "" ? font.color : "#000000",
          };
        }
      }

      // Build HTML table with per-cell styling to match the sheet
      const defaultFont = "Calibri, sans-serif";
      let htmlTable = `<table style="border-collapse: collapse; width: 100%; font-family: ${defaultFont}; table-layout: fixed;">`;

      for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
        const tag = rowIndex === 0 ? "th" : "td";
        htmlTable += "<tr>";
        const row = values[rowIndex] || [];
        for (let colIndex = 0; colIndex < colCount; colIndex++) {
          const cell = row[colIndex];
          const f = formats[rowIndex][colIndex];
          const bg =
            f.fillColor !== ""
              ? `background-color: ${toCssColor(f.fillColor)};`
              : "";
          const fontFamily = `font-family: ${f.fontName}, sans-serif;`;
          const fontSize = `font-size: ${f.fontSize}pt;`;
          const fontWeight = f.fontBold ? "font-weight: bold;" : "";
          const fontStyle = f.fontItalic ? "font-style: italic;" : "";
          const color = `color: ${toCssColor(f.fontColor)};`;
          const cellStyle = [
            "border: 1px solid #b4b4b4;",
            "padding: 4px 8px;",
            "vertical-align: bottom;",
            bg,
            fontFamily,
            fontSize,
            fontWeight,
            fontStyle,
            color,
          ]
            .filter(Boolean)
            .join(" ");
          const escaped = escapeHtml(String(cell ?? ""));
          htmlTable += `<${tag} style="${cellStyle}">${escaped}</${tag}>`;
        }
        htmlTable += "</tr>";
      }
      htmlTable += "</table>";

      const payload = {
        recipient: PAYLOAD_RECIPIENT,
        subject: `${PAYLOAD_SUBJECT}: ${new Date().toLocaleDateString()}`,
        htmlTable,
      };

      const response = await fetch(FLOW_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      if (response.ok) {
        console.log("Email sent successfully via Power Automate.");
      } else {
        const errorText = await response.text();
        console.error("Flow failed:", errorText);
        throw new Error(errorText || response.statusText);
      }
    });
  } catch (error) {
    console.error("Error in sendExcelDataToEmail:", error);
    throw error;
  }
}

/** Convert Excel color (hex or named) to a CSS-safe value. */
function toCssColor(excelColor: string): string {
  if (!excelColor) return "#000000";
  const s = excelColor.trim();
  if (/^#[0-9A-Fa-f]{6}$/.test(s)) return s;
  if (/^#[0-9A-Fa-f]{3}$/.test(s))
    return `#${s[1]}${s[1]}${s[2]}${s[2]}${s[3]}${s[3]}`;
  if (s.startsWith("rgb")) return s;
  return s;
}

function escapeHtml(text: string): string {
  const map: Record<string, string> = {
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  };
  return text.replace(/[&<>"']/g, (ch) => map[ch] ?? ch);
}
