// src/modules/cashflow/cashflow.ts

import cashflowHtml from "./cashflow.html?raw";
import { getNamedRangeValue, goToWorksheet, clearRows, setNamedRangeValue, getErrorMessage } from "../../";
import { showAboutModal } from "../../common/about";
import { CASHFLOW_ROW_START, REQUIRED_TEMPLATE_TYPE, MIN_TEMPLATE_VERSION_MAJOR, MIN_TEMPLATE_VERSION_MINOR, CURRENT_TEMPLATE_VERSION_MAJOR, CURRENT_TEMPLATE_VERSION_MINOR, CURRENT_TEMPLATE_VERSION_DATE } from "./cashflow.constants";
import { getModuleSettings, saveModuleSettings } from "../../common/userSettings";

// Module-level storage for File objects (keyed by context)
const fileStorage: { [key: string]: File | null } = {
    BEFORE: null,
    AFTER: null
};

// ---------------------------------------------
// Helper: Convert column number to Excel column letter (A, B, ..., Z, AA, AB, ...)
// Takes 1-based column number (1 = A, 2 = B, etc.)
// ---------------------------------------------
function colLetter(columnNumber: number): string {
    let result = "";
    let num = columnNumber;
    while (num > 0) {
        num--; // Convert to 0-based for calculation
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26);
    }
    return result;
}

// ---------------------------------------------
// Get File Names - Mimics VBA GetFileNames function
// Opens a file picker dialog and returns array of File objects
// ---------------------------------------------
function getFileNames(defaultFolder: string): Promise<File[]> {
    return new Promise((resolve) => {
        // Create a hidden file input element
        const fileInput = document.createElement("input");
        fileInput.type = "file";
        fileInput.multiple = true;
        fileInput.accept = ".json,.txt,.csv,.pdf,application/json,text/plain,text/csv,application/pdf";
        fileInput.style.display = "none";

        // Set initial directory if possible (browsers may ignore this)
        // Note: Browsers don't allow setting default directory for security reasons

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

        // Add to DOM, trigger click, then remove
        document.body.appendChild(fileInput);
        fileInput.click();
    });
}

// ---------------------------------------------
// Helper: Read file content as text
// ---------------------------------------------
function readFileAsText(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            resolve(e.target?.result as string);
        };
        reader.onerror = (e) => {
            reject(new Error("Failed to read file"));
        };
        reader.readAsText(file);
    });
}

// ---------------------------------------------
// Helper: Read PDF file and extract text content
// ---------------------------------------------
async function readFileAsPDF(file: File): Promise<string> {
    try {
        // Dynamically import pdfjs-dist
        // @ts-ignore - pdfjs-dist types will be available after npm install
        const pdfjsLib = await import("pdfjs-dist");
        
        // Set worker source (required for pdfjs-dist)
        // Use the version from the installed package to ensure version match
        if (typeof window !== "undefined" && pdfjsLib.GlobalWorkerOptions) {
            // Get the version from the package (pdfjs-dist exports version)
            const version = (pdfjsLib as any).version || "4.8.69";
            // Use unpkg CDN which has all versions, or fallback to cdnjs
            // Try unpkg first as it typically has all npm package versions
            pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${version}/build/pdf.worker.min.mjs`;
        }

        // Read file as array buffer
        const arrayBuffer = await file.arrayBuffer();
        
        // Load PDF document
        const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
        const pdf = await loadingTask.promise;
        
        // Extract text from all pages
        let fullText = "";
        for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
            const page = await pdf.getPage(pageNum);
            const textContent = await page.getTextContent();
            
            // Group text items by Y position to preserve line structure
            // Store items with their X position for proper ordering
            interface TextItem {
                str: string;
                x: number;
            }
            const lines: { [key: number]: TextItem[] } = {};
            
            textContent.items.forEach((item: any) => {
                if (item.str) {
                    // Get position from transform matrix
                    // transform[4] is X, transform[5] is Y
                    const xPos = item.transform ? item.transform[4] : 0;
                    const yPos = item.transform ? Math.round(item.transform[5] * 2) / 2 : 0;
                    
                    if (!lines[yPos]) {
                        lines[yPos] = [];
                    }
                    lines[yPos].push({ str: item.str, x: xPos });
                }
            });
            
            // Sort lines by Y position (descending, since PDF Y coordinates start from bottom)
            const sortedYPositions = Object.keys(lines)
                .map(Number)
                .sort((a, b) => b - a);
            
            // Build text with line breaks, sorting items within each line by X position
            const pageText = sortedYPositions
                .map(yPos => {
                    // Sort items in line by X position (left to right)
                    const sortedItems = lines[yPos].sort((a, b) => a.x - b.x);
                    return sortedItems.map(item => item.str).join(" ");
                })
                .join("\n");
            
            fullText += pageText + "\n";
        }
        
        return fullText;
    } catch (error) {
        console.error("Error reading PDF:", error);
        throw new Error(`Failed to read PDF file: ${getErrorMessage(error)}`);
    }
}

// ---------------------------------------------
// Get Schedule - Mimics VBA GetSchedule function
// Returns "1" if successful, "0" if canceled/no files
// Stores File object in module-level storage for later use
// ---------------------------------------------
export async function getSchedule(context: "BEFORE" | "AFTER"): Promise<string> {
    try {
        // Get default file path from Settings worksheet named range
        const defaultFilePath = await getNamedRangeValue("\\DefaultFilePath");
        const defaultFolder = defaultFilePath || "";

        // Get files using file picker
        const fileReports = await getFileNames(defaultFolder);

        // If the number of files selected is not exactly 1, return "0"
        if (!fileReports || fileReports.length !== 1) {
            return "0";
        }

        const file = fileReports[0];
        const fileName = file.name;
        const fileReportUpper = fileName.toUpperCase();
        
        // Store file for later use
        fileStorage[context] = file;
        
        // Check file extension and set appropriate named range
        if (fileReportUpper.includes(".CSV")) {
            await setNamedRangeValue("\\FileNGS", fileName);
            return "1";
        }
        
        if (fileReportUpper.includes(".PDF")) {
            await setNamedRangeValue("\\FileNGS", fileName);
            return "1";
        }
        
        if (fileReportUpper.includes(".JSON")) {
            if (context === "BEFORE") {
                console.log("File=" + fileReportUpper);
                await setNamedRangeValue("\\FileNGS", fileName);
            } else if (context === "AFTER") {
                await setNamedRangeValue("\\FileNCino", fileName);
            }
            return "1";
        }
    } catch (err: any) {
        console.error("Error in getSchedule:", getErrorMessage(err));
        return "0";
    }
}

// ---------------------------------------------
// Calculate NGS CSV - Placeholder for CSV processing
// ---------------------------------------------
async function calculateNGS_CSV(filePath: string, fileContent: string): Promise<void> {
    // TODO: Implement CSV processing logic
    console.log("CalculateNGS_CSV not yet implemented for file:", filePath);
    showStatusMessage("CSV processing not yet implemented.", "error");
}

// ---------------------------------------------
// Parse PDF content and extract structured data
// ---------------------------------------------
interface ParsedPDFData {
    headerFields: {
        name?: string;
        loanNo?: string;
        repay?: string;
        assocBr?: string;
        loanTerm?: string;
        contractRate?: string;
        amortizationDate?: string;
        billingRate?: string;
        installment?: string;
        maturity?: string;
    };
    amortizationRecords: Array<{
        type: string;
        eventDate: string;
        payment: number;
        interest: number;
        principal: number;
        endingBalance: number;
    }>;
    disclosure: {
        eir: number;
    };
}

function parsePDFContent(pdfText: string): ParsedPDFData {
    const lines = pdfText.split(/\r?\n/).map(line => line.trim()).filter(line => line.length > 0);
    const result: ParsedPDFData = {
        headerFields: {},
        amortizationRecords: [],
        disclosure: { eir: 0 }
    };

    // Parse header fields
    const headerText = lines.join(" ");
    
    // Extract Name (comes before "Installment")
    // Handle multiple spaces - match everything between "Name" and "Installment"
    // Use a more flexible pattern that handles commas, periods, and various characters
    const nameMatch = headerText.match(/Name\s+([^I]+?)\s+Installment/);
    if (nameMatch) {
        // Collapse multiple spaces to single space and trim
        result.headerFields.name = nameMatch[1].trim().replace(/\s+/g, " ");
    }
    
    // If first pattern didn't work, try a more permissive pattern
    if (!result.headerFields.name) {
        // Match Name followed by any characters (including commas, periods) until Installment
        const nameMatchAlt = headerText.match(/Name\s+([A-Za-z0-9,\s\.&\-']+?)\s+Installment/);
        if (nameMatchAlt) {
            result.headerFields.name = nameMatchAlt[1].trim().replace(/\s+/g, " ");
        }
    }

    // Extract Loan No (handle "Loan No" with possible multiple spaces)
    const loanNoMatch = headerText.match(/Loan\s+No\s+(\d+)/);
    if (loanNoMatch) {
        result.headerFields.loanNo = loanNoMatch[1];
    }

    // Extract Repay
    const repayMatch = headerText.match(/Repay\s+(\d+)/);
    if (repayMatch) {
        result.headerFields.repay = repayMatch[1];
    }

    // Extract Assoc/Br (can be alphanumeric like "J0403" or just numeric)
    const assocBrMatch = headerText.match(/Assoc\/Br\s+([A-Z0-9]+)/);
    if (assocBrMatch) {
        result.headerFields.assocBr = assocBrMatch[1];
    }

    // Extract Loan Term
    const loanTermMatch = headerText.match(/Loan Term\s+([^\s]+)/);
    if (loanTermMatch) {
        result.headerFields.loanTerm = loanTermMatch[1];
    }

    // Extract Contract Rate (handle multiple spaces)
    const contractRateMatch = headerText.match(/Contract\s+Rate\s+([\d.]+)/);
    if (contractRateMatch) {
        result.headerFields.contractRate = contractRateMatch[1];
    }

    // Extract Amortization Date (handle multiple spaces)
    const amortDateMatch = headerText.match(/Amortization\s+Date\s+(\d{2}\/\d{2}\/\d{4})/);
    if (amortDateMatch) {
        result.headerFields.amortizationDate = amortDateMatch[1];
    }

    // Extract Billing Rate (handle multiple spaces)
    const billingRateMatch = headerText.match(/Billing\s+Rate\s+([\d.]+)/);
    if (billingRateMatch) {
        result.headerFields.billingRate = billingRateMatch[1];
        result.disclosure.eir = parseFloat(billingRateMatch[1]);
    }

    // Extract Installment
    const installmentMatch = headerText.match(/Installment\s+([\d,]+\.\d{2})/);
    if (installmentMatch) {
        result.headerFields.installment = installmentMatch[1].replace(/,/g, "");
    }

    // Extract Maturity (handle multiple spaces, format like "FEB1", "JAN1" or date)
    // Make sure we match Maturity in the header section, not from payment schedule
    // Look for Maturity followed by spaces, then capture until we hit "Pmt" or end of header section
    // Pattern should match formats like "FEB1", "JAN1", or date formats
    const maturityMatch = headerText.match(/Maturity\s+([A-Z]{3}\d{1,2})(?:\s|$)/);
    if (maturityMatch) {
        result.headerFields.maturity = maturityMatch[1].trim();
    } else {
        // Fallback: match Maturity followed by non-space, but stop before payment table keywords
        // This ensures we don't match "Principal" from the payment schedule
        const maturityMatchAlt = headerText.match(/Maturity\s+([A-Z0-9]{2,6})(?:\s+(?:Pmt|Calc|Method|Principal|Payment|Amount|Balance|Date|#|Beginning))/);
        if (maturityMatchAlt) {
            result.headerFields.maturity = maturityMatchAlt[1].trim();
        } else {
            // Last resort: match any word after Maturity, but prefer uppercase letters + numbers
            const maturityMatchLast = headerText.match(/Maturity\s+([A-Z]{2,4}\d{0,2})/);
            if (maturityMatchLast) {
                result.headerFields.maturity = maturityMatchLast[1].trim();
            }
        }
    }

    // Find payment schedule section
    // Look for "Pmt" and "#" pattern, then "Beginning Balance"
    let beginningBalanceIndex = -1;
    let beginningBalance = 0;

    // First, find the "Beginning Balance" line
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        
        // Find "Beginning Balance" line - handle multiple spaces
        if (line.includes("Beginning") && line.includes("Balance")) {
            // Match: "Beginning   Balance   692,831.86" or similar
            const balanceMatch = line.match(/Beginning\s+Balance\s+([\d,]+\.\d{2})/);
            if (balanceMatch) {
                beginningBalance = parseFloat(balanceMatch[1].replace(/,/g, ""));
                beginningBalanceIndex = i;
                break;
            }
        }
    }

    if (beginningBalanceIndex === -1) {
        throw new Error("Could not find 'Beginning Balance' in PDF");
    }

    // Add Beginning Balance as first record (type: "CalcPmt")
    // Use empty eventDate to identify it as Beginning Balance
    // The amount will go in Column G (Net Present Value), not Column F
    result.amortizationRecords.push({
        type: "CalcPmt",
        eventDate: "", // Empty date identifies this as Beginning Balance
        payment: 0,
        interest: 0,
        principal: 0,
        endingBalance: beginningBalance // This will go in Column G (NPV)
    });

    // Parse payment rows starting after Beginning Balance
    // Payment rows have format: number date amount amount amount amount
    // Example: "1   01/01/2027   81,510.04   39,837.83   41,672.21   651,159.65"
    for (let i = beginningBalanceIndex + 1; i < lines.length; i++) {
        const line = lines[i].trim();
        
        // Stop at "Totals" row or separator lines (dashes) or "Type" line
        if (line.startsWith("Totals") || line.startsWith("Type") || line.match(/^-+$/)) {
            break;
        }

        // Skip empty lines or lines that are just dashes/separators
        if (!line || line.match(/^-+\s*$/)) {
            continue;
        }

        // Skip header lines (contain "Pmt", "Payment", "Date", etc.)
        if (line.includes("Pmt") || line.includes("Payment Date") || line.includes("Total Payment") || 
            line.includes("Interest Amount") || line.includes("Principal Amount") || line.includes("Remaining Balance") ||
            line.match(/^#\s*$/)) {
            continue;
        }

        // Match payment row: number, date, payment, interest, principal, balance
        // Pattern: number, date (MM/DD/YYYY), then 5 numbers with commas
        // Use \s+ to handle multiple spaces between fields
        const paymentMatch = line.match(/^(\d+)\s+(\d{2}\/\d{2}\/\d{4})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})/);
        
        if (paymentMatch) {
            const paymentNumber = parseInt(paymentMatch[1]);
            const eventDate = paymentMatch[2];
            const totalPayment = parseFloat(paymentMatch[3].replace(/,/g, ""));
            const interest = parseFloat(paymentMatch[4].replace(/,/g, ""));
            const principal = parseFloat(paymentMatch[5].replace(/,/g, ""));
            const endingBalance = parseFloat(paymentMatch[6].replace(/,/g, ""));

            result.amortizationRecords.push({
                type: "CalcPmt",
                eventDate: eventDate,
                payment: totalPayment,
                interest: interest,
                principal: principal,
                endingBalance: endingBalance
            });
        }
    }

    if (result.amortizationRecords.length === 0) {
        throw new Error("No payment records found in PDF");
    }

    return result;
}

// ---------------------------------------------
// Calculate NGS PDF - Parse PDF and populate Excel
// ---------------------------------------------
async function calculateNGS_PDF(filePath: string, pdfText: string): Promise<void> {
    try {
        showStatusMessage("Parsing PDF...", "ok");
        
        // Parse PDF content
        const parsedData = parsePDFContent(pdfText);

        // Set header field named ranges
        if (parsedData.headerFields.name) {
            await setNamedRangeValue("\\NGSLoanName", parsedData.headerFields.name);
        }
        if (parsedData.headerFields.loanNo) {
            await setNamedRangeValue("\\NGSLoanNumber", parsedData.headerFields.loanNo);
        }
        if (parsedData.headerFields.repay) {
            await setNamedRangeValue("\\NGSRepay", parsedData.headerFields.repay);
        }
        if (parsedData.headerFields.assocBr) {
            await setNamedRangeValue("\\NGSBranch", parsedData.headerFields.assocBr);
        }
        if (parsedData.headerFields.loanTerm) {
            await setNamedRangeValue("\\NGSLoanTerm", parsedData.headerFields.loanTerm);
        }
        if (parsedData.headerFields.contractRate) {
            await setNamedRangeValue("\\NGSContractRate", parseFloat(parsedData.headerFields.contractRate) / 100);
        }
        if (parsedData.headerFields.amortizationDate) {
            await setNamedRangeValue("\\NGSAmortizationDate", parsedData.headerFields.amortizationDate);
        }
        if (parsedData.headerFields.billingRate) {
            await setNamedRangeValue("\\NGSBillingRate", parseFloat(parsedData.headerFields.billingRate) / 100);
        }
        if (parsedData.headerFields.installment) {
            await setNamedRangeValue("\\NGSInstallment", parseFloat(parsedData.headerFields.installment));
        }
        if (parsedData.headerFields.maturity) {
            // Prefix with apostrophe to force Excel to treat as text (prevents date interpretation)
            // The apostrophe will be stored in the cell and Excel will display the value as text
            await setNamedRangeValue("\\NGSMaturity", "'" + parsedData.headerFields.maturity);
        }

        // Convert parsed data to JSON-like structure for calculateJSON
        const jsonStructure = {
            amortizationRecords: parsedData.amortizationRecords,
            disclosure: parsedData.disclosure
        };

        // Use existing calculateJSON function with the parsed data
        const jsonString = JSON.stringify(jsonStructure);
        await calculateJSON("BEFORE", jsonString);

        showStatusMessage("PDF processed successfully.", "ok");
    } catch (error) {
        console.error("Error processing PDF:", error);
        showStatusMessage(`Error processing PDF: ${getErrorMessage(error)}`, "error");
        throw error;
    }
}

// ---------------------------------------------
// Calculate JSON - Mimics VBA Calculate_JSON function
// ---------------------------------------------
export async function calculateJSON(context: "BEFORE" | "AFTER", jsonReport: string): Promise<void> {
    return Excel.run(async (excelContext) => {
        const wks = excelContext.workbook.worksheets.getItem("Schedules");
        wks.activate();

        // Determine offset and named range names based on context
        let offset: number;
        let billingRateName: string;
        let cashFlowName: string;
        let fvRateName: string;

        if (context === "BEFORE") {
            offset = 1; // Column A (VBA uses 1-based, so offset 1 = column A = Office.js index 0)
            billingRateName = "\\NGSBillingRate";
            cashFlowName = "\\CashFlowNGS";
            fvRateName = "\\FVRateNGS";
        } else {
            offset = 9; // Column I (VBA uses 1-based, so offset 9 = column I = Office.js index 8)
            billingRateName = "\\NCinoBillingRate";
            cashFlowName = "\\CashFlowNCino";
            fvRateName = "\\FVRateNCino";
        }

        // Parse JSON
        const parsedJson = JSON.parse(jsonReport);
        const payments = parsedJson.amortizationRecords;
        const paymentCount = payments.length;
        const rate = parseFloat(parsedJson.disclosure.eir);

        // Set billing rate named range (rate / 100 to convert from percentage)
        await setNamedRangeValue(billingRateName, rate / 100);

        // Convert Excel row numbers (1-based) to Office.js row indices (0-based)
        //const startRow = CASHFLOW_ROW_START - 1; // CASHFLOW_ROW_START is 13, so index is 12
        let excelRowNumber = CASHFLOW_ROW_START - 1; // Start at row 13 (index 12)
        let paymentNumber = 0;
        let hasBeginningBalance = false;

        // First pass: check if there's a Beginning Balance record
        for (let idx = 0; idx < paymentCount; idx++) {
            const record = payments[idx];
            if (record.type === "CalcPmt") {
                const isBeginningBalance = !record.eventDate || record.eventDate.trim() === "";
                if (isBeginningBalance) {
                    hasBeginningBalance = true;
                    break;
                }
            }
        }

        // Process each payment record
        for (let idx = 0; idx < paymentCount; idx++) {
            const record = payments[idx];
            if (record.type === "CalcPmt") {
                // Check if this is Beginning Balance (empty eventDate)
                const isBeginningBalance = !record.eventDate || record.eventDate.trim() === "";
                
                if (isBeginningBalance) {
                    // Beginning Balance row - special handling
                    excelRowNumber = CASHFLOW_ROW_START - 1; // Row 13 (index 12)
                    const colIndex = offset - 1;

                    // Column A: "BEGIN BALANCE" text (uppercase and bold)
                    const cellLabel = wks.getRangeByIndexes(excelRowNumber, colIndex, 1, 1);
                    cellLabel.values = [["BEGIN BALANCE"]];
                    cellLabel.format.font.bold = true;

                    // Columns B-E: Leave blank (date, payment, interest, principal)
                    // These are already empty, so no need to set them

                    // Column F (offset + 5): Remaining Balance = beginningBalance amount
                    const cellBalance = wks.getRangeByIndexes(excelRowNumber, colIndex + 5, 1, 1);
                    cellBalance.values = [[record.endingBalance]];

                    // Column G (offset + 6): Leave blank - no Net Present Value for Beginning Balance
                    // Explicitly set to empty to ensure no formula
                    const cellNPV = wks.getRangeByIndexes(excelRowNumber, colIndex + 6, 1, 1);
                    cellNPV.values = [[""]];
                } else {
                    // Regular payment row
                    paymentNumber++;
                    // Payment rows: if Beginning Balance exists, start at row 14, otherwise start at row 13
                    if (hasBeginningBalance) {
                        excelRowNumber = CASHFLOW_ROW_START + paymentNumber - 1; // Row 14 for first payment
                    } else {
                        excelRowNumber = CASHFLOW_ROW_START + paymentNumber - 2; // Row 13 for first payment
                    }

                    // Convert VBA 1-based column offset to 0-based index
                    // VBA offset 1 = Column A (index 0), but in our case offset 1 = Column B (index 1)
                    // So we use offset - 1 for the column index
                    const colIndex = offset - 1;

                    // Set payment number (column offset)
                    const cellPayment = wks.getRangeByIndexes(excelRowNumber, colIndex, 1, 1);
                    cellPayment.values = [[paymentNumber]];

                    // Set event date (column offset + 1)
                    const cellDate = wks.getRangeByIndexes(excelRowNumber, colIndex + 1, 1, 1);
                    // Strip off anything after "T" in the date string, if present
                    let dateString = record.eventDate;
                    if (typeof dateString === "string" && dateString.includes("T")) {
                        dateString = dateString.split("T")[0];
                    }
                    cellDate.values = [[dateString]];

                    // Set payment (column offset + 2)
                    const cellPaymentAmt = wks.getRangeByIndexes(excelRowNumber, colIndex + 2, 1, 1);
                    cellPaymentAmt.values = [[record.payment]];

                    // Set interest (column offset + 3)
                    const cellInterest = wks.getRangeByIndexes(excelRowNumber, colIndex + 3, 1, 1);
                    cellInterest.values = [[record.interest]];

                    // Set principal (column offset + 4)
                    const cellPrincipal = wks.getRangeByIndexes(excelRowNumber, colIndex + 4, 1, 1);
                    cellPrincipal.values = [[record.principal]];

                    // Set ending balance (column offset + 5)
                    const cellBalance = wks.getRangeByIndexes(excelRowNumber, colIndex + 5, 1, 1);
                    cellBalance.values = [[record.endingBalance]];

                    // Set formula (column offset + 6): =C[]/(1+\\FVRate)^A[]
                    // Convert lineCount (0-based) to Excel row number (1-based)
                    const excelRow = excelRowNumber + 1;
                    // Column references use 1-based column numbers
                    // Column C is offset + 2, Column A is offset
                    const colC = colLetter(offset + 2); // Column C (offset + 2)
                    const colA = colLetter(offset); // Column A (offset)
                    // Remove backslash from named range for formula (Excel formulas don't use backslash)
                    //const fvRateFormulaName = fvRateName.replace("\\", "");
                    const fvRateFormulaName = fvRateName;
                    const formula = `=${colC}${excelRow}/(1+${fvRateFormulaName})^${colA}${excelRow}`;
                    const cellFormula = wks.getRangeByIndexes(excelRowNumber, colIndex + 6, 1, 1);
                    cellFormula.formulas = [[formula]];
                }
            }
        }

        // Add TOTAL row
        // Calculate total row based on number of payments processed
        // If Beginning Balance exists: Row 13 = Beginning Balance, payments start at 14
        // If no Beginning Balance: payments start at row 13
        const totalRow = hasBeginningBalance 
            ? CASHFLOW_ROW_START + 1 + paymentNumber  // Row 26 for 12 payments with Beginning Balance
            : CASHFLOW_ROW_START + paymentNumber;     // Row 25 for 12 payments without Beginning Balance
        const rowBegin = CASHFLOW_ROW_START; // Row 13 (Beginning Balance if exists, or first payment)
        const rowEnd = totalRow - 1; // Last payment row

        // Set "TOTAL" label (uppercase and bold)
        const cellTotal = wks.getRangeByIndexes(totalRow - 1, offset - 1, 1, 1);
        cellTotal.values = [["TOTAL"]];
        cellTotal.format.font.bold = true;

        // Add SUM formulas for columns offset+2 through offset+6
        // For columns C-F (offset+2 through offset+5): include Beginning Balance row if it exists
        // For column G (offset+6, Net Present Value): exclude Beginning Balance row
        const rowBeginPayment = hasBeginningBalance ? CASHFLOW_ROW_START + 1 : CASHFLOW_ROW_START; // Row 14 if Beginning Balance exists, else row 13
        
        for (let colIdx = 2; colIdx <= 6; colIdx++) {
            const colLetterSum = colLetter(offset + colIdx);
            let formula: string;
            
            if (colIdx === 6) {
                // Column G (Net Present Value): exclude Beginning Balance row
                formula = `=SUM(${colLetterSum}${rowBeginPayment}:${colLetterSum}${rowEnd})`;
            } else {
                // Columns C-F: include Beginning Balance row if it exists
                formula = `=SUM(${colLetterSum}${rowBegin}:${colLetterSum}${rowEnd})`;
            }
            
            const cellSum = wks.getRangeByIndexes(totalRow - 1, offset - 1 + colIdx, 1, 1);
            cellSum.formulas = [[formula]];
        }

        // Get cash flow value from column offset+3 (principal column), total row
        const cashFlowCell = wks.getRangeByIndexes(totalRow - 1, offset - 1 + 3, 1, 1);
        cashFlowCell.load("values");
        await excelContext.sync();

        const cashFlowValue = cashFlowCell.values[0][0] as number;
        const formattedCashFlow = cashFlowValue.toLocaleString("en-US", {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        });

        // Set cash flow named range
        await setNamedRangeValue(cashFlowName, formattedCashFlow);

        // Format the worksheet
        await formatWorksheet("Schedules", totalRow);

        // Make total row bold
        const totalRowRange = wks.getRangeByIndexes(
            totalRow - 1,
            offset - 1,
            1,
            7 // Columns offset through offset+6
        );
        totalRowRange.format.font.bold = true;

        await excelContext.sync();
    });
}

// ---------------------------------------------
// Format Worksheet - Mimics VBA FormatWorksheet function
// ---------------------------------------------
export async function formatWorksheet(sheetName: string, rowEnd: number): Promise<void> {
    return Excel.run(async (context) => {
        const wks = context.workbook.worksheets.getItem(sheetName);
        
        // Set font for entire worksheet (only if it has data)
        // Use getUsedRangeOrNullObject to avoid errors on empty sheets
        const usedRange = wks.getUsedRangeOrNullObject();
        usedRange.load("isNullObject");
        await context.sync();
        
        if (!usedRange.isNullObject) {
            usedRange.format.font.name = "Aptos Display";
            usedRange.format.font.size = 14;
        }

        // Convert Excel row numbers (1-based) to Office.js row indices (0-based)
        const startRow = 12; // Row 13 in Excel = index 12
        const endRow = rowEnd - 1; // Convert to 0-based
        const rowCount = endRow - startRow + 1;
        
        // Validate rowCount
        if (rowCount <= 0) {
            console.warn(`Invalid row range: startRow=${startRow}, endRow=${endRow}, rowCount=${rowCount}`);
            return;
        }

        // Column A (index 0): Number format "0"
        const rangeA = wks.getRangeByIndexes(startRow, 0, rowCount, 1);
        rangeA.numberFormat = [["0"]];

        // Column B (index 1): Number format "mm/dd/yyyy"
        const rangeB = wks.getRangeByIndexes(startRow, 1, rowCount, 1);
        rangeB.numberFormat = [["mm/dd/yyyy"]];

        // Columns C-G (indices 2-6): Number format "$#,##0.00"
        const rangeCG = wks.getRangeByIndexes(startRow, 2, rowCount, 5);
        rangeCG.numberFormat = [["$#,##0.00"]];

        // Column I (index 8): Number format "0"
        const rangeI = wks.getRangeByIndexes(startRow, 8, rowCount, 1);
        rangeI.numberFormat = [["0"]];

        // Column J (index 9): Set to Excel's built-in "Short Date" style if available, else set date number format
        const rangeJ = wks.getRangeByIndexes(startRow, 9, rowCount, 1);
        rangeJ.numberFormat = [["mm/dd/yyyy"]];

        // Columns K-O (indices 10-14): Number format "$#,##0.00"
        const rangeKO = wks.getRangeByIndexes(startRow, 10, rowCount, 5);
        rangeKO.numberFormat = [["$#,##0.00"]];

        // Set font size to 18 for specific cells (0-based: A1=0,0, A5=4,0, I5=4,8)
        const cellA1 = wks.getRangeByIndexes(0, 0, 1, 1);
        cellA1.format.font.size = 18;

        const cellA5 = wks.getRangeByIndexes(4, 0, 1, 1);
        cellA5.format.font.size = 18;

        const cellI5 = wks.getRangeByIndexes(4, 8, 1, 1);
        cellI5.format.font.size = 18;

        // Set number formats for specific cells (0-based: B2=1,1, B3=2,1, B4=3,1)
        const cellB2 = wks.getRangeByIndexes(1, 1, 1, 1);
        cellB2.numberFormat = [["$#,##0.00"]];

        const cellB3 = wks.getRangeByIndexes(2, 1, 1, 1);
        cellB3.numberFormat = [["0.0000%"]];

        const cellB4 = wks.getRangeByIndexes(3, 1, 1, 1);
        cellB4.numberFormat = [["$#,##0.00"]];

        await context.sync();
    });
}

// #region Clear Cash Flow Data
// ---------------------------------------------
// CashFlow Logic (moved from taskpane.ts)
// Clears all data (both BEFORE and AFTER)
// ---------------------------------------------
/**
 * Clears the NGS and NCino cash flow areas in the Schedules worksheet,
 * matching the latest VBA logic.
 */
/**
 * Clears cash flow data for "BEFORE" or "AFTER".
 * @param contextType "BEFORE" or "AFTER"
 */
async function clearCashFlowData(contextType: "BEFORE" | "AFTER") {
    try {
        // Get the default frequency from settings (default to "Annually" if not set)
        const moduleSettings = await getModuleSettings("cashflow");
        const defaultFrequency = moduleSettings.defaultFrequency || "Annually";

        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            const wks = sheets.getItem("Schedules");

            // Properly load and sync the usedRange and rowCount
            const usedRange = wks.getUsedRangeOrNullObject();
            usedRange.load(["rowCount", "isNullObject"]);
            await context.sync();

            // Default to at least 1 row if there's nothing used (rare but possible)
            // const rowsCount =
            //     !usedRange.isNullObject && usedRange.rowCount > 0
            //         ? usedRange.rowCount
            //         : CASHFLOW_ROW_START;

            const rowsCount = 1000;

            if (contextType === "BEFORE") {
                // Clear columns A-G
                const beforeStartCol = "A";
                const beforeEndCol = "G";
                const beforeStart = beforeStartCol + CASHFLOW_ROW_START;
                const beforeEnd = beforeEndCol + rowsCount;
                const beforeRangeAddress = `${beforeStart}:${beforeEnd}`;
                const beforeRange = wks.getRange(beforeRangeAddress);
                // Clear both contents and formats in one call to ensure complete clearing
                beforeRange.clear(Excel.ClearApplyTo.all);

                // Set NGS-related named ranges to ""
                const ngsNamedRanges = [
                    "\\CashFlowNGS",
                    "\\NGSLoanName",
                    "\\NGSLoanNumber",
                    "\\NGSRepay",
                    "\\NGSBranch",
                    "\\NGSLoanTerm",
                    "\\NGSMaturity",
                    "\\NGSAmortizationDate",
                    "\\NGSInstallment",
                    "\\NGSContractRate",
                    "\\NGSBillingRate",
                    "\\FrequencyNGS"
                ];

                for (const name of ngsNamedRanges) {
                    await setNamedRangeValue(name, name === "\\FrequencyNGS" ? defaultFrequency : "");
                }
            } else if (contextType === "AFTER") {
                // Clear columns I-O
                const afterStartCol = "I";
                const afterEndCol = "O";
                const afterStart = afterStartCol + CASHFLOW_ROW_START;
                const afterEnd = afterEndCol + rowsCount;
                const afterRangeAddress = `${afterStart}:${afterEnd}`;
                const afterRange = wks.getRange(afterRangeAddress);
                afterRange.clear(Excel.ClearApplyTo.contents);
                afterRange.clear(Excel.ClearApplyTo.formats);

                // Set NCino-related named ranges to ""
                await setNamedRangeValue("\\CashFlowNCino", "");
                await setNamedRangeValue("\\FrequencyNCino", defaultFrequency);
            }
            await context.sync();
        });

        showStatusMessage(
            `CashFlow ${contextType} data cleared successfully.`,
            "ok"
        );
    } catch (err: any) {
        console.warn(
            `Failed to clear ${contextType} data: ${getErrorMessage(err)}`
        );
        showStatusMessage(
            `Error clearing ${contextType} data.`,
            "error"
        );
    }
}
// #endregion Clear Cash Flow Data

// ---------------------------------------------
// Status Message Helper (copied from taskpane.ts)
// ---------------------------------------------
function showStatusMessage(message: string, type: "ok" | "error") {
    const statusPanel = document.getElementById("status-panel")!;
    const statusText = document.getElementById("status-text")!;

    statusPanel.className = "status-panel";
    statusText.textContent = message;

    if (type === "ok") {
        statusPanel.classList.add("status-ok");
    } else {
        statusPanel.classList.add("status-bad");
    }
}

// ---------------------------------------------
// Export to PDF - Exports only the "Schedules" worksheet to PDF
// ---------------------------------------------
export async function exportToPDF(): Promise<void> {
    // Track hidden worksheets so we can restore them
    let hiddenSheetNames: string[] = [];
    
    try {
        showStatusMessage("Generating PDF...", "ok");

        // Hide all worksheets except "Schedules" before exporting
        // This ensures only the Schedules worksheet is exported to PDF
        try {
            await Excel.run(async (context) => {
                const schedulesSheet = context.workbook.worksheets.getItem("Schedules");
                schedulesSheet.activate();
                
                // Get all worksheets
                const allSheets = context.workbook.worksheets;
                allSheets.load("items/name,items/visibility");
                await context.sync();
                
                // Hide all worksheets except Schedules
                for (let i = 0; i < allSheets.items.length; i++) {
                    const sheet = allSheets.items[i];
                    if (sheet.name !== "Schedules" && sheet.visibility === Excel.SheetVisibility.visible) {
                        sheet.visibility = Excel.SheetVisibility.hidden;
                        hiddenSheetNames.push(sheet.name);
                    }
                }
                
                await context.sync();
            });
        } catch (err) {
            const errorMsg = getErrorMessage(err);
            console.error("Error preparing Schedules worksheet:", errorMsg);
            showStatusMessage(`Error: Could not find "Schedules" worksheet.`, "error");
            throw new Error(`Could not find "Schedules" worksheet: ${errorMsg}`);
        }

        // Get the workbook name for the PDF filename
        let workbookName = "Schedules";
        try {
            await Excel.run(async (context) => {
                const workbook = context.workbook;
                workbook.load("name");
                await context.sync();
                const name = workbook.name || "Workbook";
                // Remove file extension if present
                workbookName = name.replace(/\.(xlsx|xlsm|xls)$/i, "") + "_Schedules";
            });
        } catch (err) {
            console.warn("Could not get workbook name:", getErrorMessage(err));
        }

        // Use Office.js Document API to get PDF file
        // Since we activated the Schedules worksheet, only that worksheet will be exported
        return new Promise((resolve, reject) => {
            Office.context.document.getFileAsync(
                Office.FileType.Pdf,
                { sliceSize: 4194304 }, // 4 MB slices (max supported)
                (result: Office.AsyncResult<Office.File>) => {
                    if (result.status !== Office.AsyncResultStatus.Succeeded) {
                        const errorMsg = result.error?.message || "Failed to generate PDF";
                        console.error("Error getting PDF file:", errorMsg);
                        showStatusMessage(`Error: ${errorMsg}`, "error");
                        reject(new Error(errorMsg));
                        return;
                    }

                    const file = result.value;
                    const slices: Uint8Array[] = [];
                    let sliceIndex = 0;
                    const totalSlices = file.sliceCount;

                    // Function to get the next slice
                    const getNextSlice = () => {
                        file.getSliceAsync(sliceIndex, (sliceResult: Office.AsyncResult<Office.Slice>) => {
                            if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
                                const errorMsg = sliceResult.error?.message || "Failed to read PDF slice";
                                console.error("Error reading slice:", errorMsg);
                                showStatusMessage(`Error: ${errorMsg}`, "error");
                                file.closeAsync();
                                reject(new Error(errorMsg));
                                return;
                            }

                            const slice = sliceResult.value;
                            // Ensure data is Uint8Array (PDF files return byte arrays)
                            const sliceData = slice.data instanceof Uint8Array 
                                ? slice.data 
                                : new Uint8Array(slice.data);
                            slices.push(sliceData);

                            // Check if there are more slices
                            if (sliceIndex < totalSlices - 1) {
                                sliceIndex++;
                                getNextSlice();
                            } else {
                                // All slices collected, combine and download
                                file.closeAsync();

                                // Combine all slices into a single Uint8Array
                                const totalLength = slices.reduce((sum, arr) => sum + arr.length, 0);
                                const combined = new Uint8Array(totalLength);
                                let offset = 0;
                                for (const slice of slices) {
                                    combined.set(slice, offset);
                                    offset += slice.length;
                                }

                                // Create blob and download
                                const blob = new Blob([combined], { type: "application/pdf" });
                                const url = URL.createObjectURL(blob);
                                const link = document.createElement("a");
                                link.href = url;
                                link.download = `${workbookName}.pdf`;
                                document.body.appendChild(link);
                                link.click();
                                document.body.removeChild(link);
                                URL.revokeObjectURL(url);

                                showStatusMessage("PDF exported successfully.", "ok");
                                
                                // Restore hidden worksheets
                                restoreHiddenWorksheets(hiddenSheetNames).finally(() => {
                                    resolve();
                                });
                            }
                        });
                    };

                    // Start reading slices
                    getNextSlice();
                }
            );
        });
    } catch (err: any) {
        console.error("Error in exportToPDF:", getErrorMessage(err));
        showStatusMessage("Error exporting to PDF.", "error");
        
        // Restore hidden worksheets even if there was an error
        if (hiddenSheetNames.length > 0) {
            restoreHiddenWorksheets(hiddenSheetNames).catch(console.error);
        }
        
        throw err;
    }
}

// ---------------------------------------------
// Helper: Restore hidden worksheets
// ---------------------------------------------
async function restoreHiddenWorksheets(hiddenSheetNames: string[]): Promise<void> {
    if (hiddenSheetNames.length === 0) {
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            for (const sheetName of hiddenSheetNames) {
                const worksheet = context.workbook.worksheets.getItem(sheetName);
                worksheet.visibility = Excel.SheetVisibility.visible;
            }
            await context.sync();
        });
    } catch (err) {
        console.warn("Error restoring hidden worksheets:", getErrorMessage(err));
    }
}

// ---------------------------------------------
// Template Validation
// ---------------------------------------------

/**
 * Validates that the spreadsheet template is compatible with this module
 * Checks \AgTemplateType and \AgTemplateVersion named ranges
 * @returns Promise<boolean> - true if valid, false otherwise
 */
async function validateTemplate(): Promise<boolean> {
    try {
        // Check template type
        const templateType = await getNamedRangeValue("\\AgTemplateType");
        
        if (!templateType) {
            showStatusMessage(
                "Error: Named range '\\AgTemplateType' not found. This spreadsheet may not be compatible with Cash Flow Analysis.",
                "error"
            );
            return false;
        }

        if (templateType !== REQUIRED_TEMPLATE_TYPE) {
            showStatusMessage(
                `Error: Template type mismatch. Expected '${REQUIRED_TEMPLATE_TYPE}', but found '${templateType}'. This spreadsheet is not compatible with Cash Flow Analysis.`,
                "error"
            );
            return false;
        }

        // Check template version
        const templateVersion = await getNamedRangeValue("\\AgTemplateVersion");
        
        if (!templateVersion) {
            showStatusMessage(
                "Error: Named range '\\AgTemplateVersion' not found. This spreadsheet may not be compatible with Cash Flow Analysis.",
                "error"
            );
            return false;
        }

        // Parse version string (e.g., "4.7" or "4.7.0")
        const versionStr = String(templateVersion).trim();
        const versionParts = versionStr.split(".").map(part => parseInt(part, 10));
        
        if (versionParts.length < 2 || isNaN(versionParts[0]) || isNaN(versionParts[1])) {
            showStatusMessage(
                `Error: Invalid template version format '${versionStr}'. Expected format: 'major.minor' (e.g., '4.7').`,
                "error"
            );
            return false;
        }

        const templateMajor = versionParts[0];
        const templateMinor = versionParts[1];

        // Compare versions
        if (templateMajor < MIN_TEMPLATE_VERSION_MAJOR || 
            (templateMajor === MIN_TEMPLATE_VERSION_MAJOR && templateMinor < MIN_TEMPLATE_VERSION_MINOR)) {
            showStatusMessage(
                `Error: Template version '${versionStr}' is too old. Minimum required version is '${MIN_TEMPLATE_VERSION_MAJOR}.${MIN_TEMPLATE_VERSION_MINOR}'. Please update your spreadsheet template.`,
                "error"
            );
            return false;
        }

        // Validation passed
        return true;
    } catch (error) {
        console.error("Error validating template:", error);
        showStatusMessage(
            `Error validating spreadsheet template: ${getErrorMessage(error)}`,
            "error"
        );
        return false;
    }
}

// ---------------------------------------------
// Cash Flow Settings Dialog Management
// ---------------------------------------------

/**
 * Show the Cash Flow settings dialog and load current settings
 */
async function showCashFlowSettingsDialog(): Promise<void> {
    const dialog = document.getElementById("cf-settings-dialog");
    const frequencySelect = document.getElementById("cf-frequency-select") as HTMLSelectElement;
    
    if (!dialog || !frequencySelect) {
        console.error("Cash Flow settings dialog elements not found");
        return;
    }

    // Load current settings
    const moduleSettings = await getModuleSettings("cashflow");
    const currentFrequency = moduleSettings.defaultFrequency || "Annually";
    
    // Set the dropdown value
    frequencySelect.value = currentFrequency;
    
    // Show the dialog
    dialog.classList.remove("hidden");
    
    // Setup event handlers
    setupCashFlowSettingsHandlers();
}

/**
 * Hide the Cash Flow settings dialog
 */
function hideCashFlowSettingsDialog(): void {
    const dialog = document.getElementById("cf-settings-dialog");
    if (dialog) {
        dialog.classList.add("hidden");
    }
}

/**
 * Setup event handlers for the Cash Flow settings dialog
 */
function setupCashFlowSettingsHandlers(): void {
    // Close button
    const closeBtn = document.getElementById("cf-settings-close");
    if (closeBtn) {
        closeBtn.onclick = () => {
            hideCashFlowSettingsDialog();
        };
    }

    // Save button
    const saveBtn = document.getElementById("cf-settings-save");
    if (saveBtn) {
        saveBtn.onclick = async () => {
            await saveCashFlowSettings();
        };
    }
}

/**
 * Save Cash Flow settings
 */
async function saveCashFlowSettings(): Promise<void> {
    try {
        const frequencySelect = document.getElementById("cf-frequency-select") as HTMLSelectElement;
        
        if (!frequencySelect) {
            console.error("Frequency select element not found");
            return;
        }

        const selectedFrequency = frequencySelect.value;
        
        // Validate the frequency value
        const validFrequencies = ["Annually", "Semi-Annually", "Quarterly", "Monthly", "At Maturity"];
        if (!validFrequencies.includes(selectedFrequency)) {
            showStatusMessage("Invalid frequency value selected.", "error");
            return;
        }

        // Save the setting
        await saveModuleSettings("cashflow", {
            defaultFrequency: selectedFrequency
        });

        showStatusMessage("Settings saved successfully.", "ok");
        hideCashFlowSettingsDialog();
    } catch (error) {
        console.error("Error saving Cash Flow settings:", error);
        showStatusMessage("Error saving settings.", "error");
    }
}

// ---------------------------------------------
// MAIN MODULE LOADER
// ---------------------------------------------
export async function loadCashFlowModule() {
    const container = document.getElementById("module-container");
    if (!container) {
        console.error("module-container not found");
        return;
    }

    // Validate template before loading module
    const isValid = await validateTemplate();
    if (!isValid) {
        // Show error message in the container instead of loading the module
        container.innerHTML = `
            <div style="padding: 20px; text-align: center;">
                <h3 style="color: #d13438; margin-bottom: 10px;">Template Validation Failed</h3>
                <p style="color: #666;">Please check the status message above for details.</p>
                <p style="color: #666; margin-top: 10px;">Required named ranges:</p>
                <ul style="text-align: left; display: inline-block; color: #666;">
                    <li>\\AgTemplateType = "${REQUIRED_TEMPLATE_TYPE}"</li>
                    <li>\\AgTemplateVersion >= "${MIN_TEMPLATE_VERSION_MAJOR}.${MIN_TEMPLATE_VERSION_MINOR}"</li>
                </ul>
            </div>
        `;
        return;
    }

    // Inject HTML
    container.innerHTML = cashflowHtml;

    // Wire up buttons
    document.getElementById("cf-cleardata")?.addEventListener("click", () => {
        clearCashFlowData("BEFORE");
        clearCashFlowData("AFTER");
    });

    document.getElementById("cf-import-before")?.addEventListener("click", async () => {
        try {
            showStatusMessage("Selecting file...", "ok");
            
            // Get file from user
            const fileResult = await getSchedule("BEFORE");
            if (fileResult === "0") {
                showStatusMessage("No file selected.", "error");
                return;
            }

            // File is stored in fileStorage, process it
            const file = fileStorage["BEFORE"];
            if (!file) {
                showStatusMessage("File object not found.", "error");
                return;
            }

            const fileName = file.name;
            const fileUpper = fileName.toUpperCase();

            // Clear BEFORE context data before importing - ensure it completes
            showStatusMessage("Clearing existing data...", "ok");
            await clearCashFlowData("BEFORE");
            
            // Small delay to ensure clearing is fully applied
            await new Promise(resolve => setTimeout(resolve, 100));
            
            // Process based on file type
            if (fileUpper.includes(".CSV")) {
                const fileContent = await readFileAsText(file);
                await calculateNGS_CSV(fileName, fileContent);
            } else if (fileUpper.includes(".PDF")) {
                const fileContent = await readFileAsPDF(file);
                await calculateNGS_PDF(fileName, fileContent);
            } else if (fileUpper.includes(".JSON")) {
                const fileContent = await readFileAsText(file);
                await calculateJSON("BEFORE", fileContent);
                showStatusMessage("BEFORE schedule calculated successfully.", "ok");
            }
        } catch (err: any) {
            console.error("Error in Import Before:", getErrorMessage(err));
            showStatusMessage("Error importing BEFORE schedule.", "error");
        }
    });

    document.getElementById("cf-import-after")?.addEventListener("click", async () => {
        try {
            showStatusMessage("Selecting file...", "ok");
            
            // Get file from user
            const fileResult = await getSchedule("AFTER");
            if (fileResult === "0") {
                showStatusMessage("No file selected.", "error");
                return;
            }

            // File is stored in fileStorage, process it
            const file = fileStorage["AFTER"];
            if (!file) {
                showStatusMessage("File object not found.", "error");
                return;
            }

            const fileName = file.name;
            const fileUpper = fileName.toUpperCase();

            // Check file type first - AFTER only supports JSON files
            if (!fileUpper.includes(".JSON")) {
                showStatusMessage("AFTER context only supports JSON files. Please select a .JSON file.", "error");
                return;
            }

            // Clear AFTER context data before importing
            await clearCashFlowData("AFTER");
            
            // Process JSON file
            const fileContent = await readFileAsText(file);
            await calculateJSON("AFTER", fileContent);
            showStatusMessage("AFTER schedule calculated successfully.", "ok");
        } catch (err: any) {
            console.error("Error in Import After:", getErrorMessage(err));
            showStatusMessage("Error importing AFTER schedule.", "error");
        }
    });

    document.getElementById("cf-goto-schedules")?.addEventListener("click", () => {
        goToWorksheet("Schedules");
    });

    document.getElementById("cf-generate-pdf")?.addEventListener("click", async () => {
        await exportToPDF();
    });

    document.getElementById("cf-settings")?.addEventListener("click", () => {
        goToWorksheet("Settings");
    });

    document.getElementById("cf-preferences")?.addEventListener("click", async () => {
        await showCashFlowSettingsDialog();
    });

    document.getElementById("cf-about")?.addEventListener("click", async () => {
        const version = `${CURRENT_TEMPLATE_VERSION_MAJOR}.${CURRENT_TEMPLATE_VERSION_MINOR}`;
        showAboutModal("Cash Flow Analysis", version, CURRENT_TEMPLATE_VERSION_DATE);
    });
}
