/// ===============================================================
/// File: ClientReport.gs
/// Project: Template Generator – Client Report
/// Summary: Builds the Client_Report sheet dynamically based on
///          Service_Config definitions, applying Trafigura column
///          overrides when confirmed and filtering by active services.
/// ===============================================================


const INVALID_FORMULAS = [
  "=N/A",
  "=NA()",
  "N/A",
  "NA"
];

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

  // -----------------------------------------------------------------------------
  // CARGA DE CONFIGURACIÓN CON FILTRO DE SERVICIOS
  // -----------------------------------------------------------------------------
  const mappings = loadServiceMappings(trafiguraFlag, excelFlag); // Load the column configuration from Service_Config (Now from Louis Column Service Mappings), filtered by confirmed services and avoiding duplicates.

  // Retrieve the confirmed Primary Services from OperationSummary.
  const confirmedPrimaryServices = getConfirmedPrimaryServices();

  // Retrive secondary services
  const confirmedPrimarySecondaryServices = getConfirmedPrimarySecondaryServices();

  // Confirmation if only one service has choosen
  const hasSingleService =
  confirmedPrimaryServices.length === 1;

  // Create one Client_Report sheet for each confirmed Primary Service

/***
  confirmedPrimaryServices.forEach(service => {

    const formattedService =
      service.charAt(0).toUpperCase() +
      service.slice(1).toLowerCase();

    const sheetName = `Client_Report_${formattedService}`;

    if (!new_google_sheet.getSheetByName(sheetName)) {
      new_google_sheet.insertSheet(sheetName);
    }

  });
***/

  if (!hasSingleService) {

    confirmedPrimaryServices.forEach(service => {

      const formattedService =
        service.charAt(0).toUpperCase() +
        service.slice(1).toLowerCase();

      const sheetName =
        `Client_Report_${formattedService}`;

      if (!new_google_sheet.getSheetByName(sheetName)) {
        new_google_sheet.insertSheet(sheetName);
      }

    });

    

  }

  // -----------------------------------------------------------------------------
  // Populate Client_Report_<PrimaryService> sheets
  // -----------------------------------------------------------------------------

  if (!hasSingleService) {

    confirmedPrimaryServices.forEach(service => {

      const formattedService =
        service.charAt(0).toUpperCase() +
        service.slice(1).toLowerCase();

      const reportSheet =
        new_google_sheet.getSheetByName(
          `Client_Report_${formattedService}`
        );

      if (!reportSheet) {

        throw new Error(
          "[ClientReport.gs | createClientReportColumns]\n" +
          "Sheet 'Client_Report_" +
          formattedService +
          "' was not found before calling getMaxRows()."
        );

      }

      const headers = [];
      const formulas = [];

      let reportCol = 1;

      const reportMaxRows =
        reportSheet.getMaxRows();

      
      mappings.forEach(mapping => {

        // Skip columns that do not belong to the current service
        if (!mapping.primaryServices.includes(formattedService)) {
          return;
        }


        const confirmedSet =
          confirmedPrimarySecondaryServices[
            formattedService
          ];

        const serviceClassLower =
          String(mapping.serviceClass || "")
            .trim()
            .toLowerCase();

        if (
          serviceClassLower !== "prestaging" &&
          (
            !confirmedSet ||
            !confirmedSet.has(mapping.serviceClass)
          )
        ) {
          return;
        }




        const headerName =
          `${formattedService.toUpperCase()}_${mapping.columnName}`;
        
        headers.push(headerName);



        const formulaText = String(
          mapping.formula || ""
        ).trim();

        const normalizedFormula =
          formulaText.toUpperCase();

        if (
          formulaText &&
          !INVALID_FORMULAS.includes(normalizedFormula)
        ) {

          formulas.push(formulaText);

        } else {

          formulas.push("");

        }

        populateReportColumn(
          reportSheet,
          reportCol,
          mapping,
          reportMaxRows
        );

        reportCol++;

      });

      if (headers.length > 0) {

        reportSheet
          .getRange(
            1,
            1,
            1,
            headers.length
          )
          .setValues([headers])
          .setFontSize(10)
          .setFontWeight("bold")
          .setFontColor("#ffffff")
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");

      }

      if (formulas.length > 0) {

        reportSheet
          .getRange(
            2,
            1,
            1,
            formulas.length
          )
          .setFormulas([formulas]);

      }

      // -----------------------------------------------------------------------------
      // Resize Columns
      // -----------------------------------------------------------------------------

      if (reportCol > 1) {

        reportSheet.autoResizeColumns(
          1,
          reportCol - 1
        );

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

  }
  //population dynamically.

  if (!clientReportSheet) {

    throw new Error(
      "[ClientReport.gs | createClientReportColumns]\n" +
      "Client_Report sheet is null before calling getMaxRows()."
    );

  }

  const clientReportMaxRows =
    clientReportSheet.getMaxRows();

  const clientHeaders = [];
  const clientFormulas = [];

  let currentCol = 1;

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


      // Filter to avoid cross funtinality
      const confirmedSet =
        confirmedPrimarySecondaryServices[
          service
        ];

      const serviceClassLower =
        String(mapping.serviceClass || "")
          .trim()
          .toLowerCase();

      if (
        serviceClassLower !== "prestaging" &&
        (
          !confirmedSet ||
          !confirmedSet.has(mapping.serviceClass)
        )
      ) {
        return;
      }




      const headerName =
        `${service.toUpperCase()}_${mapping.columnName}`;
      
      clientHeaders.push(headerName);

      const formulaText = String(
        mapping.formula || ""
      ).trim();

      const normalizedFormula =
        formulaText.toUpperCase();

      if (
        formulaText &&
        !INVALID_FORMULAS.includes(normalizedFormula)
      ) {

        clientFormulas.push(formulaText);

      } else {

        clientFormulas.push("");

      }

      // Iterate through each mapping configuration

      // Set the header value for the current column based on the mapping's columnName.

      populateReportColumn(
        clientReportSheet,
        currentCol,
        mapping,
        clientReportMaxRows
      );

      currentCol++; // Move to the next column for the next iteration of the mapping configuration.
    });

  });

  if (clientHeaders.length > 0) {

    clientReportSheet
      .getRange(
        1,
        1,
        1,
        clientHeaders.length
      )
      .setValues([clientHeaders])
      .setFontSize(10)
      .setFontWeight("bold")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

  }

  if (clientFormulas.length > 0) {

    clientReportSheet
      .getRange(
        2,
        1,
        1,
        clientFormulas.length
      )
      .setFormulas([clientFormulas]);

  }

  // -----------------------------------------------------------------------------
  // Resizing Columns and Adding Visual Margin
  // -----------------------------------------------------------------------------
  // After all columns have been created to rezise all the columns at once

  if (currentCol > 1) {

    clientReportSheet.autoResizeColumns(
      1,
      currentCol - 1
    );


/***
    // Adding a small visual margin
    for (let c = 1; c < currentCol; c++) { // Iterate through all columns that have been populated (from column 1 to the last used column).
      clientReportSheet.setColumnWidth(
        c,
        clientReportSheet.getColumnWidth(c) + 60, // Increase the width of each column by 60 pixels to provide a visual margin for better readability.
      );
    }
***/

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
 * Populates and formats a report column.
 *
 * @param {Sheet} targetSheet
 * @param {number} targetCol
 * @param {string} headerName
 * @param {Object} mapping
 * @param {number} maxRows
 */
function populateReportColumn(
  targetSheet,
  targetCol,
  mapping,
  maxRows
) {


  if (!targetSheet) {

    throw new Error(
      "[ClientReport.gs | populateReportColumn]\n" +
      "targetSheet is null.\n" +
      "Column: " + mapping.columnName
    );

  }

  // Date / Number formatting

  if (maxRows > 2) {

    const dataRange =
      targetSheet.getRange(
        3,
        targetCol,
        maxRows - 2,
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

    } else if (
      dateType === "number"
    ) {

      dataRange.setNumberFormat(
        "0.00"
      );

    }

  }

  const calcFlag = String(
    mapping.calcFlag || ""
  )
    .trim()
    .toLowerCase();

  const isCalculated =
    calcFlag === "calculated" ||
    calcFlag === "yes";

  if (isCalculated) {

    targetSheet.getRange(
      1,
      targetCol,
      2,
      1
    ).setBackground("#e06666");

    const altRows = maxRows - 3;

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

      targetSheet
        .getRange(
          4,
          targetCol,
          altRows,
          1
        )
// Commented to reduce time
/***
        .setBackgrounds(
          colorsRedWhite
        );
***/

    }

  } else {

    targetSheet.getRange(
      1,
      targetCol,
      2,
      1
    ).setBackground("#a9a9a9");

    const altRows = maxRows - 3;

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

      targetSheet
        .getRange(
          4,
          targetCol,
          altRows,
          1
        )
// Commented to reduce time
/***
        .setBackgrounds(
          colorsGrayWhite
        );
***/

    }

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
  // Test retrival secondary services
  const confirmedPrimarySecondaryServices = getConfirmedPrimarySecondaryServices();

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

  // Activiated to see segregation
  const serviceClass = String(row[1] || "").trim();

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

  // Test secondary service
  const serviceClassLower =
  serviceClass
    .trim()
    .toLowerCase();

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
  // Test Secondary Service
  let secondaryServiceConfirmed =
    serviceClassLower === "prestaging";

  if (!secondaryServiceConfirmed) {

    secondaryServiceConfirmed =
      primaryServices.some(
        primaryService => {

          const confirmedSet =
            confirmedPrimarySecondaryServices[
              primaryService
            ];

          return (
            confirmedSet &&
            confirmedSet.has(serviceClass)
          );

        }
      );

  }

  if (!secondaryServiceConfirmed) {
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
 

  // Secondary service test
  console.log({
    column: originalColumnName,
    serviceClass: serviceClass,
    primaryServices: primaryServices
  });

  // Store the column configuration metadata for later use when
  // building and formatting the Client_Report sheet.

  mappings.push({
  //columnName: columnName,
    columnName: originalColumnName,
    serviceClass: serviceClass,
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


function debugServiceClasses() {

  const mappings =
    loadServiceMappings(0, 0);

  return mappings.map(mapping => ({
    columnName:
      mapping.columnName,

    serviceClass:
      mapping.serviceClass
  }));

}

