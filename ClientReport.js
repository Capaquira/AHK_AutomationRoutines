/// ===============================================================
/// File: ClientReport.gs
/// Project: Template Generator – Client Report
/// Summary: Builds the Client_Report sheet dynamically based on
///          Service_Config definitions, applying Trafigura column
///          overrides when confirmed and filtering by active services.
/// ===============================================================

/**
 * Builds the Client_Report sheet structure.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} new_google_sheet - Target spreadsheet instance.
 * @param {number} [trafiguraFlag=0] - 1 if Trafigura layout selected, 0 for standard.
 */

function createClientReportColumns(new_google_sheet, flag_TR, flag_EX) {
  let clientReportSheet = new_google_sheet.getSheetByName("Client_Report"); // Attempt to retrieve the existing Client_Report sheet.
  if (!clientReportSheet) {
    clientReportSheet = new_google_sheet.insertSheet("Client_Report");
  } // If the sheet doesn't exist, create it.

  // parseInt(value, 10):
  // - "value || 0" uses 0 as the default if the value is empty or undefined.
  // - "10" specifies decimal (base-10) conversion.
  const trafiguraFlag = parseInt(flag_TR || 0, 10); // 1 = Trafigura headers, 0 = Standard headers
  const excelFlag = parseInt(flag_EX || 0, 10); // 1 = Excel formulas, 0 = Google Sheets formulas

  // Log the flags for debugging
  console.log(
    "Generating Client_Report with TrafiguraFlag:",
    trafiguraFlag,
    "| Excel Formula Set:",
    excelFlag,
  );

  // -----------------------------------------------------------------------------
  // Header Formatting
  // -----------------------------------------------------------------------------

  // Set the height of the first row (header row) to 80 pixels.
  //clientReportSheet.setRowHeight(1, 80);

  // Retrieve the total number of columns currently available in the sheet.
  // This allows formatting to be applied dynamically across the entire header row,
  // regardless of future column additions or layout changes.

  const maxColumns = Math.max(
    clientReportSheet.getLastColumn(),
    1
  );


  clientReportSheet
    // Select Row 1 from Column 1 through the last available column.
    // Parameters: (startRow, startColumn, numRows, numColumns)
    // Here: Row 1, Column A, 1 row high, spanning all columns.
    .getRange(1, 1, 1, maxColumns)
    //.setWrap(true)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // -- Column Widths and Auto-Resize --
  // Column numbering starts at 1 (Column A) and continues through the last
  // available column in the sheet (maxColumns).
  for (let c = 1; c <= maxColumns; c++) {}

  // -----------------------------------------------------------------------------
  // CARGA DE CONFIGURACIÓN CON FILTRO DE SERVICIOS
  // -----------------------------------------------------------------------------
  const mappings = loadServiceMappings(trafiguraFlag, excelFlag); // Load the column configuration from Service_Config (Now from Louis Column Service Mappings), filtered by confirmed services and avoiding duplicates.

  // Retrieve the confirmed Primary Services from OperationSummary.
  const confirmedPrimaryServices = Array.from(
    getConfirmedOperationsList()
  );

  // Create one Client_Report sheet for each confirmed Primary Service

  confirmedPrimaryServices.forEach(service => {

    const formattedService =
      service.charAt(0).toUpperCase() +
      service.slice(1).toLowerCase();

    const sheetName = `Client_Report_${formattedService}`;

    if (!new_google_sheet.getSheetByName(sheetName)) {
      new_google_sheet.insertSheet(sheetName);
    }

  });

  // -----------------------------------------------------------------------------
  // Populate Client_Report_<PrimaryService> sheets
  // -----------------------------------------------------------------------------

  confirmedPrimaryServices.forEach(service => {

    const formattedService =
      service.charAt(0).toUpperCase() +
      service.slice(1).toLowerCase();

    const reportSheet =
      new_google_sheet.getSheetByName(
        `Client_Report_${formattedService}`
      );

    let reportCol = 1;

    mappings.forEach(mapping => {

      // Skip columns that do not belong to the current service
      if (!mapping.primaryServices.includes(formattedService)) {
        return;
      }

      const headerName =
        `${formattedService.toUpperCase()}_${mapping.columnName}`;

      reportSheet
        .getRange(1, reportCol)
        .setValue(headerName)
        .setFontSize(10)
        .setFontWeight("bold")
        .setFontColor("#ffffff")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");

      // Formula row
      reportSheet.getRange(2, reportCol).setValue("");

      if (mapping.formula) {

        reportSheet
          .getRange(2, reportCol)
          .setFormula(mapping.formula);

      }

      

      const reportMaxRows =
        reportSheet.getMaxRows();

      if (reportMaxRows > 2) {

        const dataRange =
          reportSheet.getRange(
            3,
            reportCol,
            reportMaxRows - 2,
            1
          );

        const dateType = String(
          mapping.dateType || ""
        )
          .trim()
          .toLowerCase();

        if (dateType === "date") {

          dataRange
            .setNumberFormat("yyyy-MM-dd")
            .setDataValidation(
              SpreadsheetApp
                .newDataValidation()
                .requireDate()
                .build()
            );

        } else if (dateType === "number") {

          dataRange.setNumberFormat("0.00");

        }

      }


      if (mapping.calcFlag === "Calculated") {

        reportSheet.getRange(
          1,
          reportCol,
          2,
          1
        ).setBackground("#e06666");

        const altRows = reportMaxRows - 3;

        if (altRows > 0) {

          const colorsRedWhite =
            Array.from(
              { length: altRows },
              (_, i) => [
                i % 2 === 0
                  ? "#fce8e6"
                  : "#ffffff"
              ]
            );

          reportSheet
            .getRange(
              4,
              reportCol,
              altRows,
              1
            )
            .setBackgrounds(
              colorsRedWhite
            );

        }

      } else {

        reportSheet.getRange(
          1,
          reportCol,
          2,
          1
        ).setBackground("#a9a9a9");

        const altRows = reportMaxRows - 3;

        if (altRows > 0) {

          const colorsGrayWhite =
            Array.from(
              { length: altRows },
              (_, i) => [
                i % 2 === 0
                  ? "#e0e0e0"
                  : "#ffffff"
              ]
            );

          reportSheet
            .getRange(
              4,
              reportCol,
              altRows,
              1
            )
            .setBackgrounds(
              colorsGrayWhite
            );

        }

      }

      reportCol++;

    });

    // -----------------------------------------------------------------------------
    // Resize Columns
    // -----------------------------------------------------------------------------

    if (reportCol > 1) {

      reportSheet.autoResizeColumns(
        1,
        reportCol - 1
      );

      for (let c = 1; c < reportCol; c++) {

        reportSheet.setColumnWidth(
          c,
          reportSheet.getColumnWidth(c) + 60
        );

      }

    }

    // -----------------------------------------------------------------------------
    // Add Filter
    // -----------------------------------------------------------------------------

    const reportLastCol =
      reportSheet.getLastColumn();

    const reportLastRow =
      reportSheet.getLastRow();

    if (reportSheet.getFilter()) {
      reportSheet.getFilter().remove();
    }

    if (
      reportLastCol > 0 &&
      reportLastRow >= 2
    ) {

      reportSheet
        .getRange(
          2,
          1,
          reportLastRow - 1,
          reportLastCol
        )
        .createFilter();

    }

  });

  const maxRows = clientReportSheet.getMaxRows(); // Get the total number of rows in the sheet to apply formatting and data population dynamically.

  let currentCol = 1; // Initialize the current column index to 1 (Column A) for populating the Client_Report sheet.

  const confirmedPrimaryServicesLower =
    confirmedPrimaryServices.map(service =>
      service.toLowerCase()
    );

  
  mappings.forEach((mapping) => {
  

    mapping.primaryServices.forEach((service) => {

      if (
        !confirmedPrimaryServicesLower.includes(
          service.toLowerCase()
        )
      ) {
        return;
      }

      const headerName =
        `${service.toUpperCase()}_${mapping.columnName}`;

      // Iterate through each mapping configuration

      // Set the header value for the current column based on the mapping's columnName.
      clientReportSheet
        .getRange(1, currentCol) // Row 1, current column
        .setValue(headerName) // Set the header text to the column name from the mapping.
        .setFontSize(10) // Set the font size
        .setFontWeight("bold") // Make the header text bold.
        .setFontColor("#ffffff")
        .setHorizontalAlignment("center") // Center-align the header text horizontally.
        .setVerticalAlignment("middle"); // Center-align the header text vertically.

      // Set the validation value for the second row of the current column based on the mapping's validation property.
      clientReportSheet.getRange(2, currentCol).setValue(""); // Set the second row to an empty string initially. This can be used for validation or additional information if needed.
      if (mapping.formula) {
        // If a formula is defined in the mapping, set it in the second row of the current column.
        clientReportSheet.getRange(2, currentCol).setFormula(mapping.formula); // Set the formula in the second row of the current column.
      }

      // Set the number format for the data rows (from row 3 to the last row)
      // based on the mapping's configured data type.
      if (maxRows > 2) {
        // Select the data range in the current column, starting at row 3.
        const dataRange = clientReportSheet.getRange(
          3, // Start row
          currentCol, // Current column
          maxRows - 2, // Number of rows from row 3 to the last row
          1, // Single column
        );
        

        // Retrieve the configured data type.
        // Using String(... || "") prevents errors if dateType is null,
        // undefined, or empty before converting to lowercase.
        const dateType = String(mapping.dateType || "")
          .trim()
          .toLowerCase();

        // Apply formatting and validation rules based on the configured data type.
        if (dateType === "date") {
          dataRange
            .setNumberFormat("yyyy-MM-dd") // Display dates using ISO format (YYYY-MM-DD)
            .setDataValidation(
              // Restrict user input to valid date values
              SpreadsheetApp.newDataValidation()
                .requireDate() // Only valid dates are allowed
                .build(),
            );
        } else if (dateType === "number") {
          dataRange.setNumberFormat("0.00"); // Display numeric values with two decimal places
        }
      }
      

      // -----------------------------------------------------------------------------
      // COLOR CODING BASED ON CALCULATION FLAG: CALCULATED VS MANUAL INPUT FIELDS
      // -----------------------------------------------------------------------------

      // If the column is marked as "Calculated", apply a red header background and alternating row colors for better visual distinction.
      if (mapping.calcFlag === "Calculated") {
        
        // Header and Formula Row Background Color
        clientReportSheet.getRange(
          1, // Row 1
          currentCol, // Current column
          2, // Rows 1 and 2
          1 // Single column
        ).setBackground("#e06666"); // Set the header background color to red for calculated fields.

        // Data Rows Background Color (Alternating Red and White)
        const altRows = maxRows - 3; // Calculate the number of rows below the header and validation rows for alternating colors.
        if (altRows > 0) { // If there are rows below the header and validation rows, apply alternating colors for better readability.
          const colorsRedWhite = Array.from({ length: altRows }, (_, i) => [
            i % 2 === 0 ? "#fce8e6" : "#ffffff", // Alternate between light red and white for calculated fields
          ]);
          clientReportSheet // Select the range of data rows in the current column and apply the alternating background colors.
            .getRange(
              4, // Start row 4 because the 3rd raw is white for all cases
              currentCol, // Current column
              altRows, // Number of rows from row 4 to the last row
              1 // Single column
            )
            .setBackgrounds(colorsRedWhite); // Set the alternating background colors for calculated fields.
        }

      } else {
        // For non-calculated fields, apply a gray header background and alternating row colors for better visual distinction.
        clientReportSheet.getRange(
          1, // Row 1
          currentCol, // Current column of the iteration
          2, // Rows 1 and 2
          1 // Single column (the current column, all column cannot be colored at one because some columns are calculated and others are not)
        ).setBackground("#a9a9a9"); // Set the header background color to gray for non-calculated fields.

        // Data Rows Background Color (Alternating Gray and White)
        const altRows = maxRows - 3; // Calculate the number of rows below the header and validation rows for alternating colors.
        if (altRows > 0) { // If there are rows below the header and validation rows, apply alternating colors for better readability.
          const colorsGrayWhite = Array.from({ length: altRows }, (_, i) => [
            i % 2 === 0 ? "#e0e0e0" : "#ffffff", // Alternate between light gray and white for non-calculated fields
          ]);
          clientReportSheet // Select the range of data rows in the current column and apply the alternating background colors.
            .getRange(
              4, // Start row 4 because the 3rd row is white for all cases
              currentCol, // Current column of the iteration
              altRows, // Number of rows from row 4 to the last row
              1 // Single column (If a value  diferente than 1 is used like 5, it will color the following 4 columns as well, which is not desired)
            )
            .setBackgrounds(colorsGrayWhite); // Set the alternating background colors for non-calculated fields.
        }
      }

    
      currentCol++; // Move to the next column for the next iteration of the mapping configuration.
    });

  });



  // -----------------------------------------------------------------------------
  // Resizing Columns and Adding Visual Margin
  // -----------------------------------------------------------------------------
  // After all columns have been created to rezise all the columns at once

  if (currentCol > 1) {

    clientReportSheet.autoResizeColumns(
      1,
      currentCol - 1
    );

    // Adding a small visual margin
    for (let c = 1; c < currentCol; c++) { // Iterate through all columns that have been populated (from column 1 to the last used column).
      clientReportSheet.setColumnWidth(
        c,
        clientReportSheet.getColumnWidth(c) + 60, // Increase the width of each column by 60 pixels to provide a visual margin for better readability.
      );
    }

  }



  // -----------------------------------------------------------------------------
  // Adding Filter to the Client_Report Sheet, below Header Row 
  // -----------------------------------------------------------------------------
  const lastColUsed = clientReportSheet.getLastColumn(); // Get the index of the last column that has data in the Client_Report sheet.
  const lastRowUsed = clientReportSheet.getLastRow(); // Get the index of the last row that has data in the Client_Report sheet.

  // Apply an automatic filter to the report if at least one column exists
  // and there are enough rows to create a valid filter range.
  if (lastColUsed > 0 && lastRowUsed >= 2) {

    if (clientReportSheet.getFilter()) {
      clientReportSheet.getFilter().remove();
    }

    clientReportSheet
      .getRange(
        2,               // Start from row 2 (formula/filter row)
        1,               // Start from column A
        lastRowUsed - 1, // Rows from row 2 to the last populated row
        lastColUsed      // All populated columns
      )
      .createFilter();
  }
}

/**
 * Reads metadata configuration from Service_Config.
 * Filters columns against confirmed services from OperationSummary
 * and prevents duplicate column names.
 *
 * @param {number} trafiguraFlag - 1 for Trafigura headers (Col B), 0 for standard headers (Col A).
 * @param {number} excelFlag - 1 for formula column M (Excel Formula), 0 for formula column L (Google Sheets Formula).
 * @returns {Array<Object>} Mappings list.
 */

function loadServiceMappings(trafiguraFlag, excelFlag) {
  const configSS = SpreadsheetApp.openByUrl(
    //"https://docs.google.com/spreadsheets/d/1FZo4hMpiJOM-cfFbC10Wb5Nug27T77wF/edit#gid=2107217982", // Old Service_Config Spreadsheet
    "https://docs.google.com/spreadsheets/d/1QqmHCH6OgZaV3kDXz8eRPpyhOffRlgIzH_nDlQ2urRI/edit#gid=2118177262", // New Service_Config Spreadsheet (Louis Column Service Mappings)
  );

  // Select the appropriate sheet based on the Trafigura flag.
  const configSheet = configSS.getSheetByName(
    trafiguraFlag === 1                     // If Trafigura flag is set to 1, use the "Traf Master Tab (Traf Mine)" sheet.
      ? "Traf Master Tab (Traf Mine)"       // Otherwise, use the "Master Tab (All other Templates)" sheet.
      : "Master Tab (All other Templates)",
  );

  // If the configuration sheet is not found, throw an error to indicate the issue.
  if (!configSheet) {
    throw new Error(
      "Sheet 'Master Tab (All other Templates) or Traf Master Tab (Traf Mine)' not found in the configuration spreadsheet.",
    );
  }

  // Last row determination
  const lastRow = configSheet.getLastRow(); // Get the last row with data in the configuration sheet to determine how many rows to read.
  if (lastRow < 2) return [];               // If there are no data rows (only header), return an empty array.

  // Data range selection
  const values = configSheet.getRange(
    2,                                      // Start from row 2 to skip the header row
    1,                                      // Start from column 1 (A)
    lastRow - 1,                            // All rows from 2 to the last one
    28                                      // All columns from A to AD
  ).getValues();

  // Confirmed services retrieval
  const confirmedServices = getConfirmedOperationsList();
  // Filter out duplicate column names
  const seenColumnNames = new Set();

  // Determine which formula column should be read from the configuration sheet.
  // Trafigura templates use column index 4, while standard templates use column index 9.
  const targetFormulaIndex = trafiguraFlag === 1 ? 4 : 9;

  // Define the Primary Service columns based on the selected configuration sheet.
  // Standard Template:
  //   Reception    = Column C (row[2])
  //   Intermediate = Column D (row[3])
  //   Loading      = Column E (row[4])
  //
  // Trafigura Template:
  //   Reception    = Column N (row[13])
  //   Loading      = Column O (row[14])
  //   Intermediate = Column P (row[15])
  const primaryServiceColumns =
    trafiguraFlag === 1
      ? {
          reception: 13,
          loading: 14,
          intermediate: 15,
        }
      : {
          reception: 2,
          intermediate: 3,
          loading: 4,
        };

  // Initialize an array to hold the final mappings after filtering and processing.
  const mappings = [];
  
  // Iterate through each row of the configuration values to process and filter them based on the confirmed services and other criteria.
  values.forEach((row) => {

  // Read the configured column name (Column A).

  /***
  let columnName = String(row[0] || "").trim();
  ***/
  const originalColumnName = String(row[0] || "").trim();

  // This part shoud be recostructed because Column B has to with Seconday Service not the Main one
  // Read the associated service class (Column B).
  //const serviceClass = String(row[1] || "").trim();

  // Determine which Primary Service(s) this field belongs to.
  const receptionFlag = String(
    row[primaryServiceColumns.reception] || ""
  ).trim().toLowerCase();

  const intermediateFlag = String(
    row[primaryServiceColumns.intermediate] || ""
  ).trim().toLowerCase();

  const loadingFlag = String(
    row[primaryServiceColumns.loading] || ""
  ).trim().toLowerCase();

  const primaryServices = [];

  if (receptionFlag === "yes") {
    primaryServices.push("Reception");
  }

  if (intermediateFlag === "yes") {
    primaryServices.push("Intermediate");
  }

  if (loadingFlag === "yes") {
    primaryServices.push("Loading");
  }

  // *** New aproach ***
  const primaryServicesLower = primaryServices.map(service =>
    service.toLowerCase()
  );
  // Service Name concadenation
  
  /***
  const primaryServiceName = (primaryServices[0] || "").toUpperCase();
  if (primaryServiceName) {
    columnName = `${primaryServiceName}_${columnName}`;
  }
  ***/
  const colKey =
  `${originalColumnName}|${primaryServices.join("|")}`
    .toLowerCase();                          

    // -------------------------------------------------------------
    // FILTRO 1: EVALUACIÓN DE SERVICIO (Columna F -> row[1])
    // -------------------------------------------------------------

    // Keep the field only when at least one of its Primary Services
    // is present in the confirmed services list.
    const isConfirmed = primaryServicesLower.some(primaryService => {

      for (const confirmed of confirmedServices) {

        const confirmedLower = String(confirmed)
          .trim()
          .toLowerCase();

        if (
          confirmedLower === primaryService ||
          confirmedLower.includes(primaryService) ||
          primaryService.includes(confirmedLower)
        ) {
          return true;
        }
      }

      return false;
    });

    if (!isConfirmed) {
      return;
    }


  // -------------------------------------------------------------
  // FILTER 2: SKIP EMPTY OR DUPLICATE COLUMN NAMES
  // -------------------------------------------------------------

  // Create a case-insensitive lookup key to ensure that column names
  // differing only by letter case are treated as duplicates.
  //const colKey = columnName.toLowerCase();

  // Skip the current configuration row if the column name is blank
  // or has already been added to the mappings collection.
  /**
  if (!columnName || seenColumnNames.has(colKey)) {
    return;
  }
  */
  if (
    !originalColumnName ||
    seenColumnNames.has(colKey)
  ) {
    return;
  }

  // Register the column name to prevent duplicate entries.
  seenColumnNames.add(colKey);

  // Store the column configuration metadata for later use when
  // building and formatting the Client_Report sheet.

  mappings.push({
  //columnName: columnName,
  columnName: originalColumnName,
  validation: String(row[20] || "").trim(),
  primaryServices: primaryServices,
  //serviceClass: serviceClass,
  dateType: String(row[18] || "").trim(),
  formula: String(row[targetFormulaIndex] || "").trim(),
  calcFlag: String(row[19] || "").trim(),
  });

  });

// Return the complete list of processed column mappings.
return mappings;
}


