// src/modules/fccg-billings/fccg-billings-report.ts
// NONAGV-COBANK Billing Notice PDF report (grouped by Company ID)
// Layout matches BillingSample.pdf: 2-line headers, column widths, currency/date formats.

import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import { Company, CompanyMap } from "./fccg-billings.types";

const WORKSHEET_AI_REPORT = "AI Report";

// AI Report column indices (0-based). Column S = 18, U = 20, V = 21. Q (Net Activity) = 16.
const AI_COL_S = 18; // companyType
const AI_COL_U = 20; // companyID
const AI_COL_V = 21; // source grouping (FCCG-..., INTERNAL-...)
const AI_COL_Q = 16; // Net Activity (for totals)
/** Company types included in the billing notice report (column S). */
const BILLING_NOTICE_COMPANY_TYPES = new Set(["NONAGV-COBANK", "COBANK", "AGV-COBANK"]);

// Report columns: [line1, line2, AI index, width pt, format: 'text'|'currency'|'date']
const REPORT_COLUMNS: { line1: string; line2: string; aiIdx: number; widthPt: number; format: "text" | "currency" | "date" }[] = [
    { line1: "Borrower", line2: "", aiIdx: 2, widthPt: 70, format: "text" },
    { line1: "Sold", line2: "Loan Number", aiIdx: 3, widthPt: 55, format: "text" },
    { line1: "FLCA", line2: "PCA", aiIdx: 4, widthPt: 35, format: "text" },
    { line1: "Interest", line2: "Rate", aiIdx: 5, widthPt: 40, format: "text" },
    { line1: "Interest", line2: "Calc Method", aiIdx: 7, widthPt: 58, format: "text" },
    { line1: "Prior Day", line2: "Part. Balance", aiIdx: 8, widthPt: 55, format: "currency" },
    { line1: "Ending", line2: "Balance", aiIdx: 9, widthPt: 55, format: "currency" },
    { line1: "Commitment", line2: "", aiIdx: 10, widthPt: 55, format: "currency" },
    { line1: "Pricing Next", line2: "Change Date", aiIdx: 11, widthPt: 52, format: "date" },
    { line1: "Ending", line2: "Accrued Interest", aiIdx: 12, widthPt: 58, format: "currency" },
    { line1: "Principal", line2: "Credit/Debit", aiIdx: 13, widthPt: 52, format: "currency" },
    { line1: "Interest Payment", line2: "Credit/Debit", aiIdx: 14, widthPt: 58, format: "currency" },
    { line1: "Net", line2: "Activity", aiIdx: 16, widthPt: 48, format: "currency" },
    { line1: "Effective", line2: "Date", aiIdx: 17, widthPt: 45, format: "date" },
    { line1: "Action For", line2: "Participating Assoc.", aiIdx: 19, widthPt: 72, format: "text" },
];

// Page layout: A4 landscape (842 x 595 points)
const PAGE_WIDTH = 842;
const PAGE_HEIGHT = 595;
// 0.25 inch border on all sides (1 inch = 72 pt)
const MARGIN = 0.25 * 72; // 18 pt
const LOGO_WIDTH = 300;
const LOGO_HEIGHT = 80;
// Logo drawn at 50% size on report
const LOGO_DISPLAY_WIDTH = 150;
const LOGO_DISPLAY_HEIGHT = 40;
const ROW_HEIGHT_MIN = 14; // minimum row height (single line)
const HEADER_ROW_HEIGHT = 28; // two lines
const FONT_SIZE = 8;
const HEADER_FONT_SIZE = 7;
const MAX_WRAP_LINES = 4; // allow rows to grow up to 4 lines
const LINE_HEIGHT_PT = 9; // baseline spacing for wrapped lines
const CELL_PAD = 2;
const TITLE_FONT_SIZE = 14;
const TITLE_LINE2_FONT_SIZE = 12;

export type ReportRow = (string | number)[];

/** Row plus metadata for grouping and totals (column V prefix, raw Net Activity). */
export interface ReportRowWithMeta {
    reportRow: ReportRow;
    netActivity: number;
    columnV: string;
}

/** Report column indices to right-justify (D=3, F=5, G=6, H=7, J=9, K=10, L=11, M=12). */
const RIGHT_JUSTIFY_COL_INDICES = new Set([3, 5, 6, 7, 9, 10, 11, 12]);

/** Build map companyID -> Company from CompanyMap (keyed by companyName). */
export function getCompanyByIDMap(companyMap: CompanyMap): Map<string, Company> {
    const byId = new Map<string, Company>();
    companyMap.forEach((company) => byId.set(company.companyID, company));
    return byId;
}

/**
 * Fetch from AI Report: rows where column S (Company Type) is NONAGV-COBANK, COBANK, or AGV-COBANK,
 * grouped by column U (Company ID). Includes column V (for FCCG/INTERNAL grouping) and raw Net Activity for totals.
 */
export async function getNONAGVCOBANKRowsGroupedByCompany(): Promise<Map<string, ReportRowWithMeta[]>> {
    const grouped = new Map<string, ReportRowWithMeta[]>();

    await Excel.run(async (context) => {
        const wks = context.workbook.worksheets.getItem(WORKSHEET_AI_REPORT);
        const usedRange = wks.getUsedRangeOrNullObject();
        usedRange.load(["rowCount", "columnCount"]);
        await context.sync();

        const lastRow = usedRange.isNullObject ? 1 : usedRange.rowCount;
        const colCount = Math.max(usedRange.isNullObject ? 0 : usedRange.columnCount, 22);
        if (lastRow <= 1) return;

        const dataRange = wks.getRangeByIndexes(1, 0, lastRow - 1, colCount);
        dataRange.load("values");
        await context.sync();

        const grid = (dataRange.values as any[][]) || [];
        for (let i = 0; i < grid.length; i++) {
            const row = grid[i] || [];
            const companyType = String(row[AI_COL_S] ?? "").trim();
            if (!BILLING_NOTICE_COMPANY_TYPES.has(companyType)) continue;

            const companyID = String(row[AI_COL_U] ?? "").trim();
            if (!companyID) continue;

            const reportRow: ReportRow = REPORT_COLUMNS.map((col) => {
                const v = row[col.aiIdx];
                if (v === undefined || v === null) return "";
                return typeof v === "number" ? v : String(v).trim();
            });
            const rawNet = row[AI_COL_Q];
            const netActivity = typeof rawNet === "number" ? rawNet : parseFloat(String(rawNet || "0").replace(/,/g, "")) || 0;
            const columnV = String(row[AI_COL_V] ?? "").trim();

            if (!grouped.has(companyID)) grouped.set(companyID, []);
            grouped.get(companyID)!.push({ reportRow, netActivity, columnV });
        }
    });

    return grouped;
}

// Net Activity column index in report (0-based)
const NET_ACTIVITY_COL_INDEX = 12;

type RenderItem =
    | { type: "data"; meta: ReportRowWithMeta }
    | { type: "total"; label: string; amount: number }
    | { type: "blank" };

/**
 * Generate the Billing Notice PDF: landscape, logo (highplains_300x80.png), two-line title,
 * 15-column table with two-line headers, FCCG/INTERNAL grouping with totals, header bg #c9e6c4.
 */
export async function generateBillingNoticePDF(
    groupedByCompany: Map<string, ReportRowWithMeta[]>,
    companyByID: Map<string, Company>,
    logoPngBytes?: Uint8Array
): Promise<Uint8Array> {
    const pdfDoc = await PDFDocument.create();
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
    const black = rgb(0, 0, 0);
    const headerBg = rgb(201 / 255, 230 / 255, 196 / 255); // #c9e6c4

    let logoImage: Awaited<ReturnType<PDFDocument["embedPng"]>> | null = null;
    if (logoPngBytes && logoPngBytes.length > 0) {
        try {
            logoImage = await pdfDoc.embedPng(logoPngBytes);
        } catch {
            logoImage = null;
        }
    }

    const contentWidth = PAGE_WIDTH - 2 * MARGIN;
    const totalColWidth = REPORT_COLUMNS.reduce((s, c) => s + c.widthPt, 0);
    const scale = totalColWidth > contentWidth ? contentWidth / totalColWidth : 1;
    const colWidths: number[] = REPORT_COLUMNS.map((c) => c.widthPt * scale);
    const colOffsets: number[] = [];
    let x = MARGIN;
    for (const w of colWidths) {
        colOffsets.push(x);
        x += w;
    }

    function addPage() {
        return pdfDoc.addPage([PAGE_WIDTH, PAGE_HEIGHT]);
    }

    function drawTitleAndHeader(
        page: ReturnType<typeof addPage>,
        companyName: string,
        yStart: number
    ): number {
        let y = yStart;

        // Logo: drawn at 50% (150x40)
        if (logoImage) {
            page.drawImage(logoImage, {
                x: MARGIN,
                y: y - LOGO_DISPLAY_HEIGHT,
                width: LOGO_DISPLAY_WIDTH,
                height: LOGO_DISPLAY_HEIGHT,
            });
        } else {
            page.drawRectangle({
                x: MARGIN,
                y: y - LOGO_DISPLAY_HEIGHT,
                width: LOGO_DISPLAY_WIDTH,
                height: LOGO_DISPLAY_HEIGHT,
                borderColor: black,
                borderWidth: 1,
            });
            page.drawText("Logo (150×40)", {
                x: MARGIN + 8,
                y: y - LOGO_DISPLAY_HEIGHT / 2 - 4,
                size: 9,
                font,
                color: black,
            });
        }

        // Title two lines: Company Name, then "Billing Notice" (centered)
        const line1W = fontBold.widthOfTextAtSize(companyName, TITLE_FONT_SIZE);
        const line2W = fontBold.widthOfTextAtSize("Billing Notice", TITLE_LINE2_FONT_SIZE);
        const titleW = Math.max(line1W, line2W);
        const titleX = MARGIN + (contentWidth - titleW) / 2;
        page.drawText(companyName, {
            x: titleX,
            y: y - LOGO_DISPLAY_HEIGHT / 2 - 4,
            size: TITLE_FONT_SIZE,
            font: fontBold,
            color: black,
        });
        page.drawText("Billing Notice", {
            x: MARGIN + (contentWidth - line2W) / 2,
            y: y - LOGO_DISPLAY_HEIGHT / 2 - 16,
            size: TITLE_LINE2_FONT_SIZE,
            font: fontBold,
            color: black,
        });

        y -= LOGO_DISPLAY_HEIGHT + 16;

        // Two-line header row: anchor from bottom so line2 clears the grid (2 lines only, not 3)
        const headerLineHeight = 9;
        const headerBottomPad = 3; // space between line2 baseline and cell bottom border
        const headerTop = y - HEADER_ROW_HEIGHT;
        const line2Baseline = headerTop + headerBottomPad;
        const line1Baseline = line2Baseline + headerLineHeight;
        REPORT_COLUMNS.forEach((col, i) => {
            const cx = colOffsets[i];
            const cw = colWidths[i];
            page.drawRectangle({
                x: cx,
                y: y - HEADER_ROW_HEIGHT,
                width: cw,
                height: HEADER_ROW_HEIGHT,
                color: headerBg,
                borderColor: black,
                borderWidth: 0.5,
            });
            const line1 = col.line1 ? truncateForCell(col.line1, cw - 4) : "";
            const line2 = col.line2 ? truncateForCell(col.line2, cw - 4) : "";
            if (line1) {
                const w1 = fontBold.widthOfTextAtSize(line1, HEADER_FONT_SIZE);
                page.drawText(line1, {
                    x: cx + (cw - w1) / 2,
                    y: line1Baseline,
                    size: HEADER_FONT_SIZE,
                    font: fontBold,
                    color: black,
                });
            }
            if (line2) {
                const w2 = fontBold.widthOfTextAtSize(line2, HEADER_FONT_SIZE);
                page.drawText(line2, {
                    x: cx + (cw - w2) / 2,
                    y: line2Baseline,
                    size: HEADER_FONT_SIZE,
                    font: fontBold,
                    color: black,
                });
            }
        });
        y -= HEADER_ROW_HEIGHT;
        return y;
    }

    function buildRenderList(rows: ReportRowWithMeta[]): RenderItem[] {
        const fccg = rows.filter((r) => r.columnV.startsWith("FCCG"));
        const internal = rows.filter((r) => r.columnV.startsWith("INTERNAL"));
        const other = rows.filter((r) => !r.columnV.startsWith("FCCG") && !r.columnV.startsWith("INTERNAL"));
        const list: RenderItem[] = [];
        fccg.forEach((m) => list.push({ type: "data", meta: m }));
        if (fccg.length > 0) {
            const sum = fccg.reduce((s, m) => s + m.netActivity, 0);
            list.push({ type: "total", label: "TOTAL FCCG BILLING", amount: sum });
        }
        if (fccg.length > 0 && internal.length > 0) list.push({ type: "blank" });
        internal.forEach((m) => list.push({ type: "data", meta: m }));
        if (internal.length > 0) {
            const sum = internal.reduce((s, m) => s + m.netActivity, 0);
            list.push({ type: "total", label: "TOTAL HIGH PLAINS DIRECT", amount: sum });
        }
        other.forEach((m) => list.push({ type: "data", meta: m }));
        return list;
    }

    function drawCellText(
        page: ReturnType<typeof addPage>,
        cx: number,
        cw: number,
        lines: string[],
        rowHeight: number,
        y: number,
        rightJustify: boolean,
        useBold = false
    ) {
        const cellFont = useBold ? fontBold : font;
        lines.forEach((line, lineIdx) => {
            const lineY = y - CELL_PAD - (lineIdx + 1) * LINE_HEIGHT_PT;
            if (lineY >= y - rowHeight + CELL_PAD) {
                const textW = cellFont.widthOfTextAtSize(line, FONT_SIZE);
                const x = rightJustify ? cx + cw - CELL_PAD - textW : cx + CELL_PAD;
                page.drawText(line, {
                    x,
                    y: lineY,
                    size: FONT_SIZE,
                    font: cellFont,
                    color: black,
                });
            }
        });
    }

    for (const [companyID, rows] of Array.from(groupedByCompany.entries())) {
        const company = companyByID.get(companyID);
        const companyName = company?.companyName ?? companyID;
        const renderList = buildRenderList(rows);

        let page = addPage();
        let y = PAGE_HEIGHT - MARGIN;
        y = drawTitleAndHeader(page, companyName, y);

        for (const item of renderList) {
            if (item.type === "blank") {
                y -= ROW_HEIGHT_MIN;
                continue;
            }

            if (item.type === "total") {
                const rowHeight = ROW_HEIGHT_MIN;
                if (y - rowHeight < MARGIN) {
                    page = addPage();
                    y = PAGE_HEIGHT - MARGIN;
                    y = drawTitleAndHeader(page, companyName, y);
                }
                const amountStr = item.amount < 0 ? `(${formatNumber(-item.amount)})` : formatNumber(item.amount);
                REPORT_COLUMNS.forEach((_col, i) => {
                    const cx = colOffsets[i];
                    const cw = colWidths[i];
                    page.drawRectangle({
                        x: cx,
                        y: y - rowHeight,
                        width: cw,
                        height: rowHeight,
                        borderColor: rgb(0.7, 0.7, 0.7),
                        borderWidth: 0.25,
                    });
                    if (i === 0) {
                        drawCellText(page, cx, cw, [item.label], rowHeight, y, false, true);
                    } else if (i === NET_ACTIVITY_COL_INDEX) {
                        drawCellText(page, cx, cw, [amountStr], rowHeight, y, true, true);
                    }
                });
                y -= rowHeight;
                continue;
            }

            const row = item.meta.reportRow;
            const cellLines: string[][] = [];
            let maxLines = 1;
            for (let i = 0; i < REPORT_COLUMNS.length; i++) {
                const col = REPORT_COLUMNS[i];
                const cw = colWidths[i];
                const maxWidthPt = cw - CELL_PAD * 2;
                const isText = col.format === "text";
                const cellVal = formatCell(row[i], col.format, isText ? 999 : maxWidthPt);
                const lines = isText
                    ? wrapText(font, cellVal, maxWidthPt, FONT_SIZE)
                    : [cellVal];
                cellLines.push(lines);
                if (lines.length > maxLines) maxLines = lines.length;
            }
            const rowHeight = Math.max(
                ROW_HEIGHT_MIN,
                CELL_PAD * 2 + maxLines * LINE_HEIGHT_PT
            );

            if (y - rowHeight < MARGIN) {
                page = addPage();
                y = PAGE_HEIGHT - MARGIN;
                y = drawTitleAndHeader(page, companyName, y);
            }

            REPORT_COLUMNS.forEach((col, i) => {
                const cx = colOffsets[i];
                const cw = colWidths[i];
                const lines = cellLines[i];
                const rightJustify = RIGHT_JUSTIFY_COL_INDICES.has(i);
                page.drawRectangle({
                    x: cx,
                    y: y - rowHeight,
                    width: cw,
                    height: rowHeight,
                    borderColor: rgb(0.7, 0.7, 0.7),
                    borderWidth: 0.25,
                });
                drawCellText(page, cx, cw, lines, rowHeight, y, rightJustify);
            });
            y -= rowHeight;
        }
    }

    return pdfDoc.save();
}

function truncateForCell(text: string, _maxWidth: number): string {
    if (text.length <= 24) return text;
    return text.slice(0, 22) + "…";
}

/** Wrap text to fit within maxWidthPt at given fontSize; return at most MAX_WRAP_LINES lines. */
function wrapText(
    font: Awaited<ReturnType<PDFDocument["embedFont"]>>,
    text: string,
    maxWidthPt: number,
    fontSize: number
): string[] {
    if (!text.trim()) return [""];
    const words = text.trim().split(/\s+/);
    const lines: string[] = [];
    let current = "";
    for (const word of words) {
        const candidate = current ? current + " " + word : word;
        const w = font.widthOfTextAtSize(candidate, fontSize);
        if (w <= maxWidthPt) {
            current = candidate;
        } else {
            if (current) {
                lines.push(current);
                if (lines.length >= MAX_WRAP_LINES) break;
            }
            const wordW = font.widthOfTextAtSize(word, fontSize);
            if (wordW <= maxWidthPt) {
                current = word;
            } else {
                // single word too long: chop to fit
                let chunk = word;
                while (chunk && font.widthOfTextAtSize(chunk, fontSize) > maxWidthPt)
                    chunk = chunk.slice(0, -1);
                lines.push(chunk + (chunk.length < word.length ? "…" : ""));
                if (lines.length >= MAX_WRAP_LINES) return lines;
                current = word.slice(chunk.length) || "";
            }
        }
    }
    if (current) lines.push(current);
    if (lines.length > MAX_WRAP_LINES) {
        const last = lines[MAX_WRAP_LINES - 1];
        let truncated = last;
        while (truncated && font.widthOfTextAtSize(truncated + "…", fontSize) > maxWidthPt)
            truncated = truncated.slice(0, -1);
        lines.length = MAX_WRAP_LINES;
        lines[MAX_WRAP_LINES - 1] = (truncated || "").replace(/…$/, "") + "…";
    }
    return lines.length ? lines : [""];
}

function parseCurrency(value: string | number): number {
    if (typeof value === "number") return value;
    const str = String(value).trim().replace(/,/g, "");
    const inParens = /^\(([^)]+)\)$/.exec(str);
    const num = parseFloat(inParens ? inParens[1] : str);
    return inParens && !isNaN(num) ? -num : num;
}

/** Format a date value as mm/dd/yyyy (handles Excel serial, ISO string, or m/d/yyyy). */
function formatDateToMMDDYYYY(value: string | number | undefined): string {
    if (value === undefined || value === null) return "";
    const v = String(value).trim();
    if (!v) return "";
    // Excel serial number (days since 1900-01-01)
    const n = parseFloat(v.replace(/,/g, ""));
    if (!isNaN(n) && n > 0) {
        const date = excelSerialToDate(n);
        if (date) {
            const m = date.getMonth() + 1;
            const d = date.getDate();
            const y = date.getFullYear();
            return `${m.toString().padStart(2, "0")}/${d.toString().padStart(2, "0")}/${y}`;
        }
    }
    // Try parsing as date string
    const d = new Date(v);
    if (!isNaN(d.getTime())) {
        const m = d.getMonth() + 1;
        const day = d.getDate();
        const y = d.getFullYear();
        return `${m.toString().padStart(2, "0")}/${day.toString().padStart(2, "0")}/${y}`;
    }
    return v;
}

function excelSerialToDate(serial: number): Date | null {
    const utc = (serial - 25569) * 86400 * 1000; // 25569 = days from 1900-01-01 to 1970-01-01
    const d = new Date(utc);
    return isNaN(d.getTime()) ? null : d;
}

function formatCell(
    value: string | number | undefined,
    format: "text" | "currency" | "date",
    maxChars: number
): string {
    if (value === undefined || value === null) return "";
    let s: string;
    if (format === "currency") {
        const num = parseCurrency(value);
        if (isNaN(num)) s = String(value);
        else if (num < 0) s = `(${formatNumber(-num)})`;
        else s = formatNumber(num);
    } else if (format === "date") {
        s = formatDateToMMDDYYYY(value);
    } else {
        s = typeof value === "number" ? String(value) : String(value);
    }
    return s.length > maxChars ? s.slice(0, maxChars - 1) + "…" : s;
}

function formatNumber(n: number): string {
    const abs = Math.abs(n);
    const int = Math.floor(abs);
    const dec = Math.round((abs - int) * 100);
    const intStr = int.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return `$${intStr}.${dec.toString().padStart(2, "0")}`;
}
