/// ===============================================================
/// File: ClientReport.gs
/// Project: Template Generator – Client Report
/// Summary: Builds the Client_Report sheet dynamically based on 
///          Service_Config definitions, applying Trafigura column 
///          overrides when confirmed.
/// ===============================================================

/**
 * Builds the Client_Report sheet structure.
 * 
 * Official Documentation References:
 * - SpreadsheetApp: https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app
 * - Sheet Class: https://developers.google.com/apps-script/reference/spreadsheet/sheet
 * - Range Class: https://developers.google.com/apps-script/reference/spreadsheet/range
 * - Filter Class: https://developers.google.com/apps-script/reference/spreadsheet/filter
 * 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} new_google_sheet - Target spreadsheet instance.
 * @param {number} [trafiguraFlag=0] - 1 if Trafigura layout selected, 0 for standard.
 */
function createClientReportColumns(new_google_sheet, trafiguraFlag) {
  // -------------------------------------------------------------
  // 1. SHEET INITIALIZATION
  // Locate or create the target output sheet named 'Client_Report'.
  // Docs: https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getsheetbynamename
  // Docs: https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#insertsheetsheetname
  // -------------------------------------------------------------
  let clientReportSheet = new_google_sheet.getSheetByName('Client_Report');
  if (!clientReportSheet) {
    clientReportSheet = new_google_sheet.insertSheet('Client_Report');
  }

  // Ensure trafiguraFlag is parsed as a clean integer (0 or 1)
  const flag = parseInt(trafiguraFlag || 0, 10);
  console.log("Generating Client_Report with TrafiguraFlag:", flag);

  // -------------------------------------------------------------
  // 2. HEADER ROW & COLUMN DIMENSION FORMATTING
  // Sets row height, horizontal & vertical alignment, text wrapping,
  // and standardized column widths.
  // Docs: https://developers.google.com/apps-script/reference/spreadsheet/sheet#setrowheightrowposition,-height
  // Docs: https://developers.google.com/apps-script/reference/spreadsheet/range#setwrapiswrap
  // Docs: https://developers.google.com/apps-script/reference/spreadsheet/sheet#setcolumnwidthcolumnposition,-width
  // -------------------------------------------------------------
  clientReportSheet.setRowHeight(1, 80);
  const maxColumns = clientReportSheet.getMaxColumns();
  clientReportSheet.getRange(1, 1, 1, maxColumns)
    .setWrap(true)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  for (let c = 1; c <= maxColumns; c++) {
    clientReportSheet.setColumnWidth(c, 200);
  }

  // -------------------------------------------------------------
  // 3. RETRIEVE CONFIGURATION & POPULATE COLUMNS
  // Fetch dynamic column schemas from Service_Config and render
  // headers, formulas, number/date formats, and alternating colors.
  // -------------------------------------------------------------
  const mappings = loadServiceMappings(flag);
  const maxRows = clientReportSheet.getMaxRows();

  let currentCol = 1;

  mappings.forEach(mapping => {
    // Write the column title to Row 1
    // Docs: https://developers.google.com/apps-script/reference/spreadsheet/range#setvaluevalue
    clientReportSheet.getRange(1, currentCol)
      .setValue(mapping.columnName)
      .setFontSize(10)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    // Clear Row 2 placeholder cell
    clientReportSheet.getRange(2, currentCol).setValue("");

    // Inject column formula if defined in the config
    // Docs: https://developers.google.com/apps-script/reference/spreadsheet/range#setformulaformula
    if (mapping.formula) {
      clientReportSheet.getRange(2, currentCol).setFormula(mapping.formula);
    }

    // Apply specific number/date formats across data rows (Row 3 downwards)
    // Docs: https://developers.google.com/apps-script/reference/spreadsheet/range#setnumberformatnumberformat
    if (maxRows > 2) {
      const dataRange = clientReportSheet.getRange(3, currentCol, maxRows - 2, 1);
      const dateType = mapping.dateType.toLowerCase();

      if (dateType === "date") {
        dataRange.setNumberFormat("yyyy-mm-dd");
      } else if (dateType === "number") {
        dataRange.setNumberFormat("0.00");
      }
    }

    // Apply color palette: Red/White for Calculated fields, Gray/White for Standard fields
    // Docs: https://developers.google.com/apps-script/reference/spreadsheet/range#setbackgroundscolor
    if (mapping.calcFlag === "Calculated") {
      clientReportSheet.getRange(1, currentCol, 2, 1).setBackground("#e06666");
      const altRows = maxRows - 3;
      if (altRows > 0) {
        const colorsRedWhite = Array.from({ length: altRows }, (_, i) => [
          (i % 2 === 0) ? '#fce8e6' : '#ffffff'
        ]);
        clientReportSheet.getRange(4, currentCol, altRows, 1).setBackgrounds(colorsRedWhite);
      }
    } else {
      clientReportSheet.getRange(1, currentCol, 2, 1).setBackground("#a9a9a9");
      const altRows = maxRows - 3;
      if (altRows > 0) {
        const colorsGrayWhite = Array.from({ length: altRows }, (_, i) => [
          (i % 2 === 0) ? '#e0e0e0' : '#ffffff'
        ]);
        clientReportSheet.getRange(4, currentCol, altRows, 1).setBackgrounds(colorsGrayWhite);
      }
    }

    currentCol++;
  });

  // -------------------------------------------------------------
  // 4. AUTO-FILTER CREATION
  // Re-creates the native spreadsheet filter across all used data rows.
  // Docs: https://developers.google.com/apps-script/reference/spreadsheet/range#createfilter
  // Docs: https://developers.google.com/apps-script/reference/spreadsheet/filter#remove
  // -------------------------------------------------------------
  const lastColUsed = clientReportSheet.getLastColumn();
  const lastRowUsed = clientReportSheet.getLastRow();

  if (lastColUsed > 0 && lastRowUsed >= 2) {
    const existingFilter = clientReportSheet.getFilter();
    if (existingFilter) {
      existingFilter.remove();
    }
    clientReportSheet.getRange(2, 1, lastRowUsed - 1, lastColUsed).createFilter();
  }
}

/**
 * Reads metadata configuration from the central Service_Config spreadsheet.
 * Selects Column B headers when Trafigura flag is active, defaulting to Column A.
 * 
 * Official Documentation References:
 * - SpreadsheetApp.openByUrl: https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app#openbyurlurl
 * - Range.getValues: https://developers.google.com/apps-script/reference/spreadsheet/range#getvalues
 * 
 * @param {number} trafiguraFlag - 1 for Trafigura column headers (Col B), 0 for standard (Col A).
 * @returns {Array<Object>} Mapped column specifications.
 */
function loadServiceMappings(trafiguraFlag) {
  // -------------------------------------------------------------
  // 1. OPEN CONFIGURATION WORKBOOK & SHEET
  // Access external central configuration document.
  // -------------------------------------------------------------
  const configSS = SpreadsheetApp.openByUrl(
    "https://docs.google.com/spreadsheets/d/1FZo4hMpiJOM-cfFbC10Wb5Nug27T77wF/edit#gid=2107217982"
  );
  const configSheet = configSS.getSheetByName('Service_Config');

  if (!configSheet) {
    throw new Error("Sheet 'Service_Config' not found in the configuration spreadsheet.");
  }

  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) return [];

  // -------------------------------------------------------------
  // 2. FETCH CELL DATA RANGE (A2:L)
  // Extract values in bulk into memory to avoid individual API calls.
  // -------------------------------------------------------------
  const values = configSheet.getRange(2, 1, lastRow - 1, 12).getValues();

  // -------------------------------------------------------------
  // 3. PARSE & MAP ATTRIBUTES DYNAMICALLY
  // Target index: 1 (Column B) for Trafigura, 0 (Column A) for Standard.
  // -------------------------------------------------------------
  const targetColIndex = (trafiguraFlag === 1) ? 1 : 0;
  const mappings = [];

  values.forEach(row => {
    // Select column header by flag; fallback to Column A if Column B is blank
    let columnName = String(row[targetColIndex] || "").trim();
    if (trafiguraFlag === 1 && !columnName) {
      columnName = String(row[0] || "").trim();
    }

    const serviceClass = String(row[5] || "").trim();
    const dateType     = String(row[7] || "").trim();
    const calcFlag     = String(row[8] || "").trim();
    const validation   = String(row[9] || "").trim();
    const formula      = String(row[11] || "").trim();

    if (columnName) {
      mappings.push({
        columnName: columnName,
        validation: validation,
        serviceClass: serviceClass,
        dateType: dateType,
        formula: formula,
        calcFlag: calcFlag
      });
    }
  });

  return mappings;
}