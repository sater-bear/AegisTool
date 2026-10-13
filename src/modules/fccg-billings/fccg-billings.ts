// src/modules/fccg-billings/fccg-billings.ts

import fccgBillingsHtml from "./fccg-billings.html?raw";
import { goToWorksheet, clearRows, getNamedRangeValue, getErrorMessage, autoFitAllColumns, sendExcelDataToEmail } from "../../";
import { showAboutModal } from "../../common/about";
import { Company, CompanyMap } from "./fccg-billings.types";
import {
    getNONAGVCOBANKRowsGroupedByCompany,
    getCompanyByIDMap,
    generateBillingNoticePDF,
} from "./fccg-billings-report";
import * as XLSX from "xlsx";

// Worksheet name constants
const WORKSHEET_PART_SOLD_REPORT = "AI Report";
const WORKSHEET_BILLING_SUMMARY = "Billing Summary";
const WORKSHEET_SYSTEM = "System";
const FLOW_URL = "https://defaultfcb6fca8f3c7458b9a80211f5ce84b.61.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/d2347d905f904fcab6a7111a25b2d8e0/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=VrMbrCseOphfZrttLPKm-3whsAo8O58_z7NbXoL3VN0";
const RECIPIENT = "your-email@company.com";
const SUBJECT = "Excel Data Export";

// Table name constant
const TABLE_COMPANY = "\\tblCompany";

// ---------------------------------------------
// MAIN MODULE LOADER
// ---------------------------------------------
export async function loadFCCGBillingsModule() {
    const container = document.getElementById("module-container");
    if (!container) {
        console.error("module-container not found");
        return;
    }

    // Inject HTML
    container.innerHTML = fccgBillingsHtml;

    // Wire up buttons - PREPARE section
    document.getElementById("fccg-cleardata")?.addEventListener("click", async () => {
        await clearAllData();
    });

    // Wire up buttons - BILLINGS section
    document.getElementById("fccg-billings-goto-report")?.addEventListener("click", async () => {
        await goToWorksheet(WORKSHEET_PART_SOLD_REPORT);
    });

    document.getElementById("fccg-billings-import-reports")?.addEventListener("click", async () => {
        await importParticipationsSoldWorkbooks();
    });

    document.getElementById("fccg-billings-goto-summary")?.addEventListener("click", async () => {
        await goToWorksheet(WORKSHEET_BILLING_SUMMARY);
    });

    document.getElementById("fccg-billings-summarize")?.addEventListener("click", async () => {
        // TODO: Implement Summarize Billings functionality
        console.log("Summarize Billings clicked");
    });

    document.getElementById("fccg-billings-run")?.addEventListener("click", async () => {
        // TODO: Implement Run Billings functionality
        console.log("Run Billings clicked");
    });

    document.getElementById("fccg-billings-send")?.addEventListener("click", async () => {
        // TODO: Implement Send Billings functionality
        console.log("Send Billings clicked");
    });

    document.getElementById("fccg-billings-billing-notice-pdf")?.addEventListener("click", async () => {
        await runBillingNoticePDF();
    });

    document.getElementById("fccg-billings-send-excel-to-email")?.addEventListener("click", async () => {
        try {
            showStatusMessage("Sending data to email...", "ok");
            await sendExcelDataToEmail();
            showStatusMessage("Data sent to email successfully.", "ok");
        } catch (error) {
            console.error("Send to email error:", getErrorMessage(error));
            showStatusMessage("Error sending data to email.", "error");
        }
    });

    // Wire up buttons - FEES section
    document.getElementById("fccg-fees-goto-report")?.addEventListener("click", () => {
        // TODO: Implement Go To Report functionality
        console.log("Fees Go To Report clicked");
    });

    document.getElementById("fccg-fees-send-calculator")?.addEventListener("click", async () => {
        // TODO: Implement Send to Calculator functionality
        console.log("Send to Calculator clicked");
    });

    document.getElementById("fccg-fees-goto-data-entry")?.addEventListener("click", () => {
        // TODO: Implement Go To Data Entry functionality
        console.log("Go To Data Entry clicked");
    });

    document.getElementById("fccg-fees-calculate")?.addEventListener("click", async () => {
        // TODO: Implement Calculate Fee functionality
        console.log("Calculate Fee clicked");
    });

    document.getElementById("fccg-fees-send-jv")?.addEventListener("click", async () => {
        // TODO: Implement Send to JV functionality
        console.log("Send to JV clicked");
    });

    // Wire up buttons - JV section
    document.getElementById("fccg-jv-goto")?.addEventListener("click", () => {
        // TODO: Implement Go To JV functionality
        console.log("Go To JV clicked");
    });

    document.getElementById("fccg-jv-run-notices")?.addEventListener("click", async () => {
        // TODO: Implement Run Fee Notices functionality
        console.log("Run Fee Notices clicked");
    });

    document.getElementById("fccg-jv-send-notices")?.addEventListener("click", async () => {
        // TODO: Implement Send Fee Notices functionality
        console.log("Send Fee Notices clicked");
    });

    document.getElementById("fccg-jv-clear")?.addEventListener("click", async () => {
        // TODO: Implement Clear JV functionality
        console.log("Clear JV clicked");
    });

    // Wire up buttons - DNT/WIRES section
    document.getElementById("fccg-dnt-goto-fccg")?.addEventListener("click", () => {
        // TODO: Implement Go To DNT-FCCG functionality
        console.log("Go To DNT-FCCG clicked");
    });

    document.getElementById("fccg-dnt-goto-ip")?.addEventListener("click", () => {
        // TODO: Implement Go To DNT-IP functionality
        console.log("Go To DNT-IP clicked");
    });

    document.getElementById("fccg-dnt-send-cobank")?.addEventListener("click", async () => {
        // TODO: Implement Send To CoBank functionality
        console.log("Send To CoBank clicked");
    });

    document.getElementById("fccg-dnt-send-fee-wires")?.addEventListener("click", async () => {
        // TODO: Implement Send Fee Wires functionality
        console.log("Send Fee Wires clicked");
    });

    document.getElementById("fccg-dnt-send-billing-wires")?.addEventListener("click", async () => {
        // TODO: Implement Send Billing Wires functionality
        console.log("Send Billing Wires clicked");
    });

    // Wire up buttons - SETTINGS section
    document.getElementById("fccg-about")?.addEventListener("click", async () => {
        // TODO: Get version from constants when created
        showAboutModal("FCCG - Billings & Fees", "1.0.0", new Date().toLocaleDateString());
    });
}

// ---------------------------------------------
// Clear All Data Function
// ---------------------------------------------
/**
 * Clears all data from Part. Sold Report and Billing Summary worksheets
 * Equivalent to VBA ClearAllData() function
 */
async function clearAllData(): Promise<void> {
    try {
        showStatusMessage("Clearing all data...", "ok");

        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            
            // Get the worksheets
            const wksPartSold = worksheets.getItem(WORKSHEET_PART_SOLD_REPORT);
            const wksSummary = worksheets.getItem(WORKSHEET_BILLING_SUMMARY);

            // Note: Office.js AutoFilter API doesn't have a direct way to remove/clear AutoFilter
            // The VBA code turns off AutoFilter, but in Office.js this isn't directly supported
            // We'll skip this step as it's not critical - clearing the rows will work regardless
            // If AutoFilter is active, it will remain but won't interfere with data clearing

            // Clear outlines from Billing Summary
            // Note: Office.js doesn't have a direct ClearOutline method, 
            // but we can clear the outline by clearing the range and reformatting
            // For now, we'll just clear the rows as the main functionality

            // Clear rows starting from row 2
            // VBA calls ClearRows(wksSummary, 2) twice and ClearRows(wks, 2) once
            await clearRows(wksSummary, 2, false);
            await clearRows(wksPartSold, 2, false);
            await clearRows(wksSummary, 2, false);

            // Select Part. Sold Report worksheet and select cell A2
            wksPartSold.activate();
            const cellA2 = wksPartSold.getRange("A2");
            cellA2.select();

            await context.sync();
        });

        showStatusMessage("All data cleared successfully.", "ok");
    } catch (error) {
        console.error("Error clearing all data:", getErrorMessage(error));
        showStatusMessage("Error clearing all data.", "error");
    }
}

// ---------------------------------------------
// Get File Names - For selecting Excel files
// ---------------------------------------------
function getExcelFileNames(): Promise<File[]> {
    return new Promise((resolve) => {
        const fileInput = document.createElement("input");
        fileInput.type = "file";
        fileInput.multiple = true;
        fileInput.accept = ".xlsx,.xlsm,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel";
        fileInput.style.display = "none";

        fileInput.addEventListener("change", (event) => {
            const target = event.target as HTMLInputElement;
            const files = target.files;
            
            if (files && files.length > 0) {
                const fileArray: File[] = [];
                for (let i = 0; i < files.length; i++) {
                    fileArray.push(files[i]);
                }
                document.body.removeChild(fileInput);
                resolve(fileArray);
            } else {
                document.body.removeChild(fileInput);
                resolve([]);
            }
        });

        document.body.appendChild(fileInput);
        fileInput.click();
    });
}

// ---------------------------------------------
// Import Participations Sold Workbooks
// ---------------------------------------------
/**
 * Imports 3 Excel files and concatenates them into Part. Sold Report
 * Equivalent to VBA ImportParticipationsSoldWorkbooks() function
 */
async function importParticipationsSoldWorkbooks(): Promise<void> {
    try {
        showStatusMessage("Selecting Excel files...", "ok");

        // Get 3 Excel files from user
        const files = await getExcelFileNames();
        
        if (!files || files.length === 0) {
            showStatusMessage("No files selected.", "error");
            return;
        }

        // if (files.length !== 3) {
        //     showStatusMessage("Please select exactly 3 Excel files.", "error");
        //     return;
        // }

        // Get company map
        showStatusMessage("Loading company map...", "ok");
        const companyMap = await getCompanyMap();

        // Get effective date (we'll need to get this from a named range or input)
        const effectiveDate = new Date().toLocaleDateString();

        // Process each file in order: whole content from row 2 onward; first at row 2, rest appended
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            const wks = worksheets.getItem(WORKSHEET_PART_SOLD_REPORT);
            const rngStart = wks.getRange("A1");
            rngStart.select();
            await context.sync();

            for (const file of files) {
                await importParticipationsSoldWorkbook(wks, file.name, file);
            }
            await context.sync();
        });

        // Process the imported data
        showStatusMessage("Processing imported data...", "ok");
        await processImportedData(companyMap, effectiveDate);

        // Format the worksheet (font, column R as Short Date)
        await formatPartSoldReportWorksheet();

        // Count total rows imported
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            const wks = worksheets.getItem(WORKSHEET_PART_SOLD_REPORT);
            const usedRange = wks.getUsedRangeOrNullObject();
            usedRange.load("rowCount");
            await context.sync();
            
            const totalRows = usedRange.isNullObject ? 0 : usedRange.rowCount;
            showStatusMessage(`Import completed successfully. Total rows: ${totalRows}`, "ok");
        });
    } catch (error) {
        console.error("Error importing reports:", getErrorMessage(error));
        showStatusMessage("Error importing reports.", "error");
    }
}

// ---------------------------------------------
// Import Participations Sold Workbook
// ---------------------------------------------
/**
 * Imports one workbook: whole sheet content from row 2 onward.
 * First file writes starting at destination row 2; subsequent files append at the end.
 */
async function importParticipationsSoldWorkbook(
    wks: Excel.Worksheet,
    fileName: string,
    file: File
): Promise<void> {
    try {
        showStatusMessage(`Importing: ${fileName}...`, "ok");

        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        const allRows: any[][] = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            defval: "",
            raw: false
        });

        // Whole content starting at row 2 (skip source row 1)
        let data = allRows.length <= 1 ? [] : allRows.slice(1);

        // Remove rows where Excel column E (index 4) does not contain "FLCA" or "PCA"
        const EXCEL_COL_E_INDEX = 3; // E = 4th column, 0-based index 3
        data = data.filter((row) => {
            const flcaPcaValue = String((row || [])[EXCEL_COL_E_INDEX] ?? "").toUpperCase();
            return flcaPcaValue == "FLCA" || flcaPcaValue == "PCA";
        });

        if (!data.length) {
            console.warn(`No data from row 2 onward in ${fileName} (after filtering by FLCA/PCA in column E)`);
            return;
        }

        // Column A = Type (from file name); source column A → AI Report column B, B → C, etc.
        const fileNameUpper = fileName.toUpperCase();
        const sourceType =
            fileNameUpper.startsWith("COBANK") ? "COBANK" :
            fileNameUpper.startsWith("FCCG") ? "FCCG" :
            fileNameUpper.startsWith("INTERNAL") ? "INTERNAL" : "";

        const usedRange = wks.getUsedRangeOrNullObject();
        usedRange.load("rowCount");
        await wks.context.sync();

        // First file → start at destination row 2 (index 1); more files → append after last row
        const startRowIndex = (usedRange.isNullObject || usedRange.rowCount === 0)
            ? 1
            : usedRange.rowCount;

        const sourceColCount = Math.max(...data.map(row => (row || []).length), 1);
        const insertBlankBeforeG = sourceType === "FCCG" || sourceType === "INTERNAL";
        const colCount = sourceColCount + 1 + (insertBlankBeforeG ? 1 : 0); // +1 for Type; +1 for blank before G when FCCG/INTERNAL
        const EXCEL_COL_G_INDEX = 6; // G = 7th column, 0-based index 6
        const values: any[][] = [];
        for (let i = 0; i < data.length; i++) {
            const sourceRow = (data[i] || []).slice();
            while (sourceRow.length < sourceColCount) sourceRow.push("");
            const row: any[] = [sourceType, ...sourceRow];
            if (insertBlankBeforeG) {
                row.splice(EXCEL_COL_G_INDEX, 0, ""); // insert blank column before G
            }
            values.push(row);
        }
        const rowCount = values.length;

        const targetRange = wks.getRangeByIndexes(startRowIndex, 0, rowCount, colCount);
        targetRange.values = values;
        await wks.context.sync();

        console.log(`Imported ${rowCount} rows from ${fileName}`);
    } catch (error) {
        console.error(`Error importing ${fileName}:`, getErrorMessage(error));
        throw error;
    }
}

// ---------------------------------------------
// Process Imported Data
// ---------------------------------------------
/**
 * Processes the imported data after all files are imported
 * Equivalent to the second part of VBA ImportParticipationsSoldWorkbooks()
 */
async function processImportedData(companyMap: CompanyMap, effectiveDate: string): Promise<void> {
    try {
        showStatusMessage("Processing imported data...", "ok");

        const missingCompanies: string[] = [];

        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            const wks = worksheets.getItem(WORKSHEET_PART_SOLD_REPORT);

            const usedRange = wks.getUsedRangeOrNullObject();
            usedRange.load(["rowCount", "columnCount"]);
            await context.sync();

            const lastRow = usedRange.isNullObject ? 1 : usedRange.rowCount;
            const colCount = Math.max(usedRange.isNullObject ? 0 : usedRange.columnCount, 22);
            if (lastRow <= 1) return;

            // Load entire data block in one round-trip (rows 2 to lastRow, columns A–V)
            const dataRange = wks.getRangeByIndexes(1, 0, lastRow - 1, colCount);
            dataRange.load("values");
            await context.sync();

            const grid: any[][] = (dataRange.values as any[][]).map(row => {
                const r = (row || []).slice();
                while (r.length < 22) r.push("");
                return r;
            });

            // Process in memory (no per-row sync)
            for (let rowIdx = 0; rowIdx < grid.length; rowIdx++) {
                const row = grid[rowIdx];
                const borrowerName = String(row[2] || "").trim();

                row[16] = ""; // clear Q

                if (borrowerName === "" || borrowerName === "Primary Borrower") continue;

                let source = String(row[0] || "").trim();
                let billingCompany = String(row[1] || "").trim();
                const flcaPca = String(row[4] || "").trim();
                const principal = parseFloat(String(row[13] || "0").replace(/,/g, ""));
                const interest = parseFloat(String(row[14] || "0").replace(/,/g, ""));

                if (billingCompany === "Farm Credit Services of Colusa-Glenn, FLCA") {
                    billingCompany = "Farm Credit Services of Colusa-Glenn, FL";
                } else if (billingCompany === "Farm Credit Services of Colusa-Glenn, ACA") {
                    billingCompany = "Farm Credit Services of Colusa-Glenn, AC";
                }

                const company = companyMap.get(billingCompany);

                if (company) {
                    const netActivity = ((principal * -1) - interest) * -1;
                    let processReportLine = netActivity !== 0;
                    if (principal !== 0 && interest !== 0 && Math.abs(principal) === Math.abs(interest)) {
                        processReportLine = true;
                        row[15] = "Y"; // P
                    }
                    if (processReportLine) {
                        row[16] = netActivity;   // Q
                        row[17] = effectiveDate; // R
                        row[18] = company.companyType; // S
                        row[19] = netActivity < 0 ? "CREDIT" : "DEBIT"; // T
                        row[20] = company.companyID; // U
                        row[21] = `${source}-${company.companyID}-${flcaPca}`; // V
                    }
                } else {
                    const rowKey = `${rowIdx + 2} ==> ${billingCompany}`;
                    if (!missingCompanies.includes(rowKey)) missingCompanies.push(rowKey);
                }
            }

            dataRange.values = grid;
            await context.sync();
        });

        if (missingCompanies.length > 0) {
            showStatusMessage("The Following Companies Are Not Mapped: " + missingCompanies.join("; "), "error");
        }
    } catch (error) {
        console.error("Error processing imported data:", getErrorMessage(error));
        throw error;
    }
}

// ---------------------------------------------
// Format Part Sold Report Worksheet
// ---------------------------------------------
/**
 * Formats the Part. Sold Report worksheet: Calibri 11pt, no bold for data rows;
 * column R (Effective Date) as Excel Short Date. Row 1 (headers) is not changed.
 */
async function formatPartSoldReportWorksheet(): Promise<void> {
    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            const wks = worksheets.getItem(WORKSHEET_PART_SOLD_REPORT);

            const usedRange = wks.getUsedRangeOrNullObject();
            usedRange.load(["rowCount", "columnCount"]);
            await context.sync();

            const lastRow = usedRange.isNullObject ? 1 : usedRange.rowCount;
            const colCount = usedRange.isNullObject ? 0 : usedRange.columnCount;
            if (lastRow <= 1 || colCount <= 0) return;

            // Format from Excel row 2 onward; do not format row 0 (Excel row 1)
            const dataRowCount = lastRow - 1; // rows 2 to lastRow (0-based index 1 to lastRow-1)
            const dataRange = wks.getRangeByIndexes(1, 0, dataRowCount, colCount);
            dataRange.format.font.name = "Segoe UI";
            dataRange.format.font.size = 11;
            dataRange.format.font.bold = false;

            if (colCount > 17) {
                const columnR = wks.getRangeByIndexes(1, 17, dataRowCount, 1);
                columnR.numberFormat = [["mm/dd/yyyy"]];
            }
            await context.sync();
        });

        await autoFitAllColumns(WORKSHEET_PART_SOLD_REPORT);
    } catch (error) {
        console.error("Error formatting Part. Sold Report:", getErrorMessage(error));
        throw error;
    }
}

// ---------------------------------------------
// Billing Notice PDF (Company Type NONAGV-COBANK / COBANK / AGV-COBANK, grouped by Company ID)
// ---------------------------------------------
/**
 * Builds PDF from AI Report rows where column S is NONAGV-COBANK, COBANK, or AGV-COBANK, grouped by column U (Company ID).
 * Landscape, logo placeholder, title "Company Name Billing Notice", 15-column table,
 * title/headers repeated each page, page break between companies. Triggers download.
 */
async function runBillingNoticePDF(): Promise<void> {
    try {
        showStatusMessage("Preparing Billing Notice PDF...", "ok");
        const companyMap = await getCompanyMap();
        const companyByID = getCompanyByIDMap(companyMap);
        const grouped = await getNONAGVCOBANKRowsGroupedByCompany();
        if (grouped.size === 0) {
            showStatusMessage("No rows with Company Type (column S) NONAGV-COBANK, COBANK, or AGV-COBANK found on AI Report.", "error");
            return;
        }
        showStatusMessage("Generating PDF...", "ok");
        let logoPngBytes: Uint8Array | undefined;
        try {
            const logoUrl = new URL("/assets/highplains_300x80.png", window.location.href).href;
            const logoRes = await fetch(logoUrl);
            if (logoRes.ok) {
                const ab = await logoRes.arrayBuffer();
                logoPngBytes = new Uint8Array(ab);
            }
        } catch {
            // use no logo (placeholder will be drawn)
        }
        const pdfBytes = await generateBillingNoticePDF(grouped, companyByID, logoPngBytes);
        const blob = new Blob([pdfBytes as BlobPart], { type: "application/pdf" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = "BillingNotice.pdf";
        link.click();
        URL.revokeObjectURL(url);
        showStatusMessage("Billing Notice PDF downloaded.", "ok");
    } catch (error) {
        console.error("Error generating Billing Notice PDF:", getErrorMessage(error));
        showStatusMessage("Error generating Billing Notice PDF.", "error");
    }
}

// ---------------------------------------------
// Get Company Map Function
// ---------------------------------------------
/**
 * Gets the company map from the System worksheet table
 * Equivalent to VBA GetCompanyMap() function
 * @returns Promise<CompanyMap> - Map of companies keyed by CompanyName
 */
export async function getCompanyMap(): Promise<CompanyMap> {
    const companyMap: CompanyMap = new Map();

    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            const wksSystem = worksheets.getItem(WORKSHEET_SYSTEM);

            // Get the table
            const tbl = wksSystem.tables.getItem(TABLE_COMPANY);
            const tblRange = tbl.getRange();
            
            // Load table data
            tblRange.load(["rowCount", "columnCount", "values"]);
            await context.sync();

            // Get named range values for email bodies
            const emailBodyBilling = await getNamedRangeValue("\\EmailBodyBilling");
            const emailBodyWireFee = await getNamedRangeValue("\\EmailBodyWireFee");

            // Iterate through rows starting at row 2 (index 1 in 0-based)
            // VBA: For intIdx = 2 To tbl.Range.Rows.Count
            for (let rowIdx = 1; rowIdx < tblRange.rowCount; rowIdx++) {
                const row = tblRange.values[rowIdx];
                
                // Get CompanyName from column 1 (index 0)
                const companyName = String(row[0] || "").trim();
                
                // Skip if CompanyName is "0" or empty
                if (companyName === "" || companyName === "0") {
                    continue;
                }

                // Create company object
                const company: Company = {
                    companyName: companyName,
                    companyID: String(row[1] || "").trim(), // Column 2
                    companyType: String(row[2] || "").trim(), // Column 3
                    glPayableAccount: "34026", // Hardcoded as in VBA
                    wireScope: String(row[3] || "").trim(), // Column 4
                    glAccount_FCCG: String(row[4] || "").trim(), // Column 5
                    accountingUnit_FCCG_FLCA: String(row[5] || "").trim(), // Column 6
                    accountingUnit_FCCG_PCA: String(row[6] || "").trim(), // Column 7
                    glAccount_IP: String(row[7] || "").trim(), // Column 8
                    accountingUnit_IP_FLCA: String(row[8] || "").trim(), // Column 9
                    accountingUnit_IP_PCA: String(row[9] || "").trim(), // Column 10
                    includeJV: String(row[10] || "").trim(), // Column 11
                    worksheetNameBillings: String(row[11] || "").trim(), // Column 12
                    worksheetNameWireBilling: String(row[12] || "").trim(), // Column 13
                    worksheetNameFees: String(row[13] || "").trim(), // Column 14
                    worksheetNameWireFee: String(row[14] || "").trim(), // Column 15
                    emailAddressFees_FCCG: String(row[15] || "").trim(), // Column 16
                    emailBodyFee_FCCG: String(row[16] || "").trim(), // Column 17
                    emailAddressBilling_FCCG: String(row[17] || "").trim(), // Column 18
                    emailBodyBilling: emailBodyBilling || "", // From named range
                    emailAddressWire: "accounting@highplainsfarmcredit.com", // Hardcoded as in VBA
                    emailBodyWireFee: emailBodyWireFee || "" // From named range
                };

                // Add to map keyed by CompanyName
                companyMap.set(companyName, company);
            }
        });

        return companyMap;
    } catch (error) {
        console.error("Error getting company map:", getErrorMessage(error));
        showStatusMessage("Error getting company map.", "error");
        throw error;
    }
}

// ---------------------------------------------
// Status Message Helper
// ---------------------------------------------
function showStatusMessage(message: string, type: "ok" | "error") {
    const statusPanel = document.getElementById("status-panel")!;
    const statusText = document.getElementById("status-text")!;

    statusPanel.className = "status-panel"; // reset classes
    statusText.textContent = message;

    if (type === "ok") {
        statusPanel.classList.add("status-ok");
    } else {
        statusPanel.classList.add("status-bad");
    }
}
