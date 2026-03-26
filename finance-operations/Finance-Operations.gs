/**
 * Configuration Constants
 */
const CONFIG = {
  STAGING_SHEET: "SYS_Import_Staging",
  LEDGER_SHEET: "Import_List",
  DELIMITER: "|",
};

/**
 * Initializes the custom menu upon opening the spreadsheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Financial Ops")
    .addItem("Import CAMT V8 CSV", "showImportDialog")
    .addToUi();
}

/**
 * Serves the HTML modal dialog.
 */
function showImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile("index")
    .setWidth(400)
    .setHeight(300)
    .setTitle("CSV Ingestion Engine");
  SpreadsheetApp.getUi().showModalDialog(html, "Import CAMT V8 CSV");
}

/**
 * Ensures required infrastructure exists. Creates sheets if missing.
 */
function ensureInfrastructureExists_(ss) {
  [CONFIG.STAGING_SHEET, CONFIG.LEDGER_SHEET].forEach((sheetName) => {
    if (!ss.getSheetByName(sheetName)) {
      ss.insertSheet(sheetName);
    }
  });
}

/**
 * Normalizes cell values for cryptographic hashing.
 * Strips whitespace, invisible chars, and currency symbols to prevent visual formatting drift.
 */
function normalizeForHash_(value) {
  return String(value).replace(/[\s€]/g, "").trim();
}

/**
 * Core ingestion and deduplication engine. Called from client-side JS.
 * @param {string} csvText - The raw CSV string from the file upload.
 * @return {Object} Payload containing success status and operation metrics.
 */
function processCsvUpload(csvText) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureInfrastructureExists_(ss);

    const stagingSheet = ss.getSheetByName(CONFIG.STAGING_SHEET);
    const ledgerSheet = ss.getSheetByName(CONFIG.LEDGER_SHEET);

    // -------------------------------------------------------------
    // 1. Sanitize and Parse Data (Handles BOMs, Double Quotes, and Semicolons)
    // -------------------------------------------------------------
    let cleanText = csvText.replace(/^\uFEFF/, "").trim();
    if (cleanText.startsWith('""')) {
      cleanText = cleanText.substring(1);
    }

    let csvData;
    try {
      csvData = Utilities.parseCsv(cleanText, ";");
    } catch (parseError) {
      return {
        error:
          "Native parsing failed. The CSV structure or quoting is severely malformed.",
      };
    }

    if (!csvData || csvData.length === 0) {
      return { error: "File is empty or invalid CSV format." };
    }

    // -------------------------------------------------------------
    // 2. Dump to Volatile Staging
    // -------------------------------------------------------------
    stagingSheet.clearContents();
    stagingSheet
      .getRange(1, 1, csvData.length, csvData[0].length)
      .setValues(csvData);

    // Force Sheets to apply locale formatting before re-reading
    SpreadsheetApp.flush();

    // Re-read staging through getDisplayValues() for consistent hashing.
    // This ensures CSV data matches ledger's getDisplayValues() output.
    const stagingData = stagingSheet
      .getRange(1, 1, csvData.length, csvData[0].length)
      .getDisplayValues();

    // -------------------------------------------------------------
    // 3. Deduplication Engine Setup
    // -------------------------------------------------------------
    const ledgerRange = ledgerSheet.getDataRange();

    // CRITICAL: getDisplayValues() prevents JS Date/Number coercion, ensuring 1:1 string match
    const ledgerData = ledgerRange.getDisplayValues();
    const existingKeys = new Set();
    const hasLedgerData =
      ledgerData.length > 0 && ledgerData[0].join("").trim() !== "";

    // Column count safety guard — abort if CSV and ledger column counts differ
    if (hasLedgerData) {
      const ledgerCols = ledgerData[0].length;
      const csvCols = csvData[0].length;
      if (csvCols !== ledgerCols) {
        return {
          error: `Column mismatch: CSV has ${csvCols} columns but Ledger has ${ledgerCols}. Aborting to prevent data corruption.`,
        };
      }
    }

    // Build Composite Keys for existing ledger
    if (hasLedgerData) {
      for (let i = 0; i < ledgerData.length; i++) {
        const hashKey = ledgerData[i]
          .map(normalizeForHash_)
          .join(CONFIG.DELIMITER);
        existingKeys.add(hashKey);
      }
    }

    // -------------------------------------------------------------
    // 4. Ingestion & Comparison Phase
    // -------------------------------------------------------------
    const newRecords = [];
    let duplicatesIgnored = 0;

    // CRITICAL FIX: If ledger already has headers, skip index 0 (the CSV header)
    // to prevent prepending duplicate headers into the data body.
    const startIndex = hasLedgerData ? 1 : 0;

    for (let i = startIndex; i < stagingData.length; i++) {
      const rowKey = stagingData[i].map(normalizeForHash_).join(CONFIG.DELIMITER);

      if (existingKeys.has(rowKey)) {
        duplicatesIgnored++;
      } else {
        newRecords.push(stagingData[i]);
        existingKeys.add(rowKey); // Prevent intra-file duplication
      }
    }

    // -------------------------------------------------------------
    // 5. Commit Phase (Insert After Headers)
    // -------------------------------------------------------------
    if (newRecords.length > 0) {
      const numCols = newRecords[0].length;

      if (!hasLedgerData) {
        // Condition A: Ledger is empty. Write starting at Row 1 (Includes headers).
        ledgerSheet
          .getRange(1, 1, newRecords.length, numCols)
          .setValues(newRecords);
        ledgerSheet.setFrozenRows(1); // Freeze headers for UX
      } else {
        // Condition B: Ledger has data. Insert new rows right after the header row
        // so the newest imports always appear at the top.
        ledgerSheet.insertRowsAfter(1, newRecords.length);
        ledgerSheet
          .getRange(2, 1, newRecords.length, numCols)
          .setValues(newRecords);
      }
    }

    // Force recalculation so dependent ARRAYFORMULA sheets pick up new data immediately.
    SpreadsheetApp.flush();

    // -------------------------------------------------------------
    // 6. Cleanup Staging — delete the sheet entirely to declutter
    // -------------------------------------------------------------
    ss.deleteSheet(stagingSheet);

    return {
      success: true,
      total: csvData.length - startIndex, // Calculate actual data rows processed
      ignored: duplicatesIgnored,
      added: newRecords.length,
    };
  } catch (error) {
    return { error: error.toString() };
  }
}


/**
 * DIAGNOSTIC TOOL: Run this manually from the editor if duplicates still occur.
 * It compares Row 2 of Staging vs Row 2 of Ledger to pinpoint formatting drift.
 */
function runHashDiagnostic() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = ss.getSheetByName(CONFIG.LEDGER_SHEET);
  const stagingSheet = ss.getSheetByName(CONFIG.STAGING_SHEET);

  if (!ledgerSheet || !stagingSheet) return Logger.log("Missing sheets.");

  const ledgerRow = ledgerSheet
    .getRange(2, 1, 1, ledgerSheet.getLastColumn())
    .getDisplayValues()[0];
  const stagingRow = stagingSheet
    .getRange(2, 1, 1, stagingSheet.getLastColumn())
    .getDisplayValues()[0];

  Logger.log("=== DIAGNOSTIC REPORT ===");
  Logger.log(
    "Ledger Hash : " + ledgerRow.map(normalizeForHash_).join(CONFIG.DELIMITER),
  );
  Logger.log(
    "Staging Hash: " + stagingRow.map(normalizeForHash_).join(CONFIG.DELIMITER),
  );
  Logger.log("-------------------------");

  const maxCols = Math.max(ledgerRow.length, stagingRow.length);
  for (let i = 0; i < maxCols; i++) {
    const lVal = normalizeForHash_(ledgerRow[i] || "");
    const sVal = normalizeForHash_(stagingRow[i] || "");
    if (lVal !== sVal) {
      Logger.log(
        `MISMATCH at Index [${i}]: Ledger="${lVal}" vs Staging="${sVal}"`,
      );
    }
  }
}
