/**
 * @OnlyCurrentDoc
 * This script automates stock table updates and sorting within a Google Sheet,
 * typically triggered by user edits.
 */

// --- Constants (Centralized Configuration) ---
const COLUMN_INDICES = {
  STOCK_SYMBOL: 1, // Column A
  ENTRY_PRICE: 2,  // Column B
  YESTERDAY_PRICE: 3, // Column C
  PROFIT_LOSS: 4,     // Column D
  PROFIT_LOSS_PERCENT: 5, // Column E
  SORT_COLUMN: 9    // Column I for sorting
};

const TABLE_START_ROW = 5; // Assuming your stock data table starts on row 5, not including headers.

// --- Helper Functions (Re-using some from previous simplification, if applicable) ---

/**
 * Gets the active spreadsheet.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} The active spreadsheet.
 */
function getActiveSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Gets the active sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The active sheet.
 */
function getActiveSheet() {
  return getActiveSpreadsheet().getActiveSheet();
}

/**
 * Finds the 1-based row number of a cell containing specific text.
 * Optimized for finding header rows within the first few rows.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search.
 * @param {string} searchText The text to find.
 * @param {number} [searchLimitRows=10] The number of rows to search within (e.g., 10 for headers).
 * @returns {number | null} The 1-based row number if found, otherwise null.
 */
function findHeaderRow(sheet, searchText, searchLimitRows = 10) {
  const values = sheet.getRange(1, 1, Math.min(searchLimitRows, sheet.getLastRow()), sheet.getLastColumn()).getValues();
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (String(values[r][c]).trim() === searchText.trim()) {
        return r + 1; // Return 1-based row index
      }
    }
  }
  return null; // Not found
}

/**
 * Finds the 1-based column number of a cell containing specific text within a given row.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search.
 * @param {string} searchText The text to find.
 * @param {number} row The 1-based row number to search within.
 * @returns {number | null} The 1-based column number if found, otherwise null.
 */
function findHeaderColumn(sheet, searchText, row) {
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (let c = 0; c < values.length; c++) {
    if (String(values[c]).trim() === searchText.trim()) {
      return c + 1; // Return 1-based column index
    }
  }
  return null; // Not found
}


// --- Main Functions (Refined Logic) ---

/**
 * An onEdit trigger function that updates the stock table whenever a cell is edited.
 * This function is automatically called by Google Sheets when an edit occurs.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object.
 */
function onEdit(e) {
  Logger.log("onEdit triggered. Updating stocks table.");
  // Optional: Add logging or specific checks based on the edited range.
  // Example: if (e.range.getColumn() === COLUMN_INDICES.STOCK_SYMBOL) { UpdateStocksTable(); }
  UpdateStocksTable();
}

/**
 * Sorts the main stock data table by a predefined column.
 * Assumes the table starts at TABLE_START_ROW and covers up to Column I (9).
 */
function sortTable() {
  Logger.log("Sorting the stock data table.");
  const sheet = getActiveSheet();

  // Determine the actual last row of data in the table.
  // Using Column A to find the last row of data, assuming it's always populated.
  const lastDataRow = sheet.getRange(sheet.getMaxRows(), COLUMN_INDICES.STOCK_SYMBOL)
                           .getNextDataCell(SpreadsheetApp.Direction.UP)
                           .getRow();

  // The sort range starts from TABLE_START_ROW, extends to the last data row,
  // and covers from Column A (1) to the sort column (9).
  // Subtracting TABLE_START_ROW to get the number of rows from the start of the table.
  const sortRange = sheet.getRange(TABLE_START_ROW, 1, lastDataRow - TABLE_START_ROW + 1, COLUMN_INDICES.SORT_COLUMN);

  // Sort by the defined sort column (COLUMN_INDICES.SORT_COLUMN), in descending order.
  sortRange.sort([{column: COLUMN_INDICES.SORT_COLUMN, ascending: false}]);
  Logger.log("Stock data table sorted.");
}


/**
 * Updates various financial formulas for stock data in the current sheet.
 * This includes current price, yesterday's price, profit/loss, and percentage.
 * It also handles newly added stock symbols by populating their formulas.
 */
function UpdateStocksTable() {
  Logger.log("Starting stock table update.");
  const sheet = getActiveSheet();

  // Determine the actual last row with a stock symbol (Column A)
  const lastTickerRow = sheet.getRange(sheet.getMaxRows(), COLUMN_INDICES.STOCK_SYMBOL)
                             .getNextDataCell(SpreadsheetApp.Direction.UP)
                             .getRow();

  // Find the header row (e.g., "Stock Symbol")
  const headerRow = findHeaderRow(sheet, "Stock Symbol");
  if (!headerRow) {
    Logger.log("Header 'Stock Symbol' not found. Cannot update table.");
    return;
  }

  // --- Identify New Rows for Initialization ---
  // Find the last row with data in Column B (Entry Price) and Column C (Yesterday Price)
  // This helps identify rows where data/formulas haven't been placed yet.
  const lastEntryPriceRow = sheet.getRange(sheet.getMaxRows(), COLUMN_INDICES.ENTRY_PRICE)
                                 .getNextDataCell(SpreadsheetApp.Direction.UP)
                                 .getRow();
  const lastYesterdayPriceRow = sheet.getRange(sheet.getMaxRows(), COLUMN_INDICES.YESTERDAY_PRICE)
                                     .getNextDataCell(SpreadsheetApp.Direction.UP)
                                     .getRow();

  // Start initialization from the row after the last data in Columns B/C, up to the last ticker row.
  const initStartRow = Math.max(headerRow + 1, lastEntryPriceRow + 1, lastYesterdayPriceRow + 1);
  const rowsToInitialize = lastTickerRow - initStartRow + 1;

  if (rowsToInitialize > 0) {
    Logger.log(`Initializing formulas for ${rowsToInitialize} new rows.`);
    // Batch set values/formulas for new rows
    const tickerRange = sheet.getRange(initStartRow, COLUMN_INDICES.STOCK_SYMBOL, rowsToInitialize, 1);
    const tickers = tickerRange.getValues().map(row => row[0]);

    // Prepare arrays for formulas to be set in a single batch operation
    const entryPriceFormulas = [];
    const yesterdayPriceFormulas = [];

    tickers.forEach((ticker, i) => {
      const currentRow = initStartRow + i;
      // Current Price: This was originally "<Enter Entry Price>", then GOOGLEFINANCE.
      // If it's a manual entry, it should be left blank or pre-filled.
      // Assuming you want it to be a manual entry initially for the client.
      entryPriceFormulas.push(["<Enter Entry Price>"]); // Or empty string: [""]
      yesterdayPriceFormulas.push([`=GOOGLEFINANCE(A${currentRow})`]); // Google Finance for yesterday's price
    });

    // Batch update entry price column (Column B)
    sheet.getRange(initStartRow, COLUMN_INDICES.ENTRY_PRICE, rowsToInitialize, 1).setValues(entryPriceFormulas);

    // Batch update yesterday's price column (Column C)
    sheet.getRange(initStartRow, COLUMN_INDICES.YESTERDAY_PRICE, rowsToInitialize, 1).setFormulas(yesterdayPriceFormulas);
  } else {
    Logger.log("No new rows to initialize.");
  }


  // --- Update Existing Rows' Formulas (Profit/Loss, Percentage) ---
  // This loop runs for all rows from after the header to the last ticker row.
  const updateStartRow = headerRow + 1;
  const rowsToUpdate = lastTickerRow - updateStartRow + 1;

  if (rowsToUpdate <= 0) {
    Logger.log("No existing rows to update with profit/loss formulas.");
    return;
  }

  Logger.log(`Updating profit/loss formulas for ${rowsToUpdate} rows.`);

  // Prepare arrays for formulas to be set in a single batch operation
  const profitLossFormulas = [];
  const profitLossPercentFormulas = [];

  for (let i = 0; i < rowsToUpdate; i++) {
    const currentRow = updateStartRow + i;

    // Profit/Loss (Column D): =IFERROR(C - B, 0.00)
    profitLossFormulas.push([`=IFERROR(C${currentRow}-B${currentRow},0.00)`]);

    // Profit/Loss Percentage (Column E): =IFERROR(D / B, 0.00%)
    profitLossPercentFormulas.push([`=IFERROR(D${currentRow}/B${currentRow},0.00%)`]);
  }

  // Batch update Profit/Loss column (Column D)
  sheet.getRange(updateStartRow, COLUMN_INDICES.PROFIT_LOSS, rowsToUpdate, 1)
       .setFormulas(profitLossFormulas)
       .setNumberFormat("#,##0.00"); // Apply number format once

  // Batch update Profit/Loss Percentage column (Column E)
  sheet.getRange(updateStartRow, COLUMN_INDICES.PROFIT_LOSS_PERCENT, rowsToUpdate, 1)
       .setFormulas(profitLossPercentFormulas)
       .setNumberFormat("0.00%"); // Apply number format once

  Logger.log("Stock table update complete.");

  // Important: The original code had commented-out sections for font coloring
  // based on profit/loss or 'Buy/Sell/Hold' status.
  // If this functionality is desired, it should be implemented in a separate
  // dedicated function (e.g., `applyConditionalFormatting()`)
  // and called explicitly if needed, or better, use native Google Sheet
  // Conditional Formatting rules which are more performant.
}

// --- Minor / Unused Functions (Consider removal or refactoring) ---

// The findCellRow and findCellColumn functions provided in the original code
// are inefficient as they iterate through all cells.
// The `findHeaderRow` and `findHeaderColumn` helpers above provide a more
// targeted and efficient way to find headers. If the original functions
// are not used elsewhere or for other specific, large data searches,
// they can be removed.
/*
function findCellRow(strKeyword) { ... }
function findCellColumn(strKeyword) { ... }
*/

// Original onEdit commented-out line:
// // var range = e.range;
// // range.setNote('Last modified: ' + new Date());
// This functionality is simple and can be re-added directly to onEdit if desired.
// However, setting notes on every edit can clutter a sheet quickly.