Write the complete code (`Code.gs` and `index.html`) based on the architectural requirements below. Prioritize execution speed, memory efficiency (O(1) lookups), and batch operations to avoid Google Sheets API rate limits and timeouts.

### 1. Sheet Infrastructure

The script must interact with (and create if missing) two specific sheets in the active spreadsheet:

- `SYS_Import_Staging`: A volatile worksheet used solely for dumping the parsed raw CSV array before processing.
- `Master_Ledger`: The persistent database of all historical transactions.

### 2. UI & Ingestion Flow

- Create an `onOpen` function that builds a Custom Menu: "Financial Ops" -> "Import CAMT V8 CSV".
- Clicking the menu opens a modal dialog (`index.html`) using `HtmlService`.
- The HTML UI should be clean, modern, and contain a file input accepting `.csv` and an "Upload" button.
- Pass the file content to a server-side function. Parse the CSV text into a 2D array using `Utilities.parseCsv()`.
- Dump this raw array directly into `SYS_Import_Staging` (clear any existing data in this sheet first).

### 3. Deduplication Engine (Strict 100% Match)

This is the critical path. Do NOT use row-by-row `Range.getValues()` comparisons. Implement a Cryptographic or Composite Key hashing mechanism using native JavaScript `Set` objects to ensure O(1) lookup complexity.

- **Step A:** Read all existing data from `Master_Ledger` in one batch query.
- **Step B:** Generate a unique signature (Composite Key) for every existing row by concatenating all column values in that row into a single string (using a specific delimiter like `|`). Store all existing keys in a JavaScript `Set`.
- **Step C:** Read the freshly imported data from `SYS_Import_Staging`.
- **Step D:** Iterate through the new data. Generate the exact same Composite Key format for each row.
  - If the key exists in the `Set`, it is a duplicate -> discard the row.
  - If the key does not exist -> push the row to a `newRecords` array AND add the new key to the `Set` (to prevent duplicate entries from within the same upload file).
- **Note:** Treat all CSV data as immutable strings during this comparison phase. Do not format dates or numbers.

### 4. Commit Phase

- If `newRecords` has length > 0, append the array to the bottom of `Master_Ledger` using a single `Range.setValues()` batch operation.
- Clear the `SYS_Import_Staging` sheet upon successful commit.
- Surface a standard `SpreadsheetApp.getUi().alert()` detailing: "Import successful. [Total] rows processed. [Ignored] duplicates skipped. [Added] new records appended."

Ensure the code is fully modular, well-commented, handles potential errors (like empty files or incorrect formats), and is ready for immediate deployment.
