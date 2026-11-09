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
  if (!clientReportSheet) {clientReportSheet = new_google_sheet.insertSheet("Client_Report");} // If the sheet doesn't exist, create it.

  // parseInt(value, 10):
  // - "value || 0" uses 0 as the default if the value is empty or undefined.
  // - "10" specifies decimal (base-10) conversion.
  const trafiguraFlag = parseInt(flag_TR || 0, 10); // 1 = Trafigura headers, 0 = Standard headers
  const excelFlag = parseInt(flag_EX || 0, 10);     // 1 = Excel formulas, 0 = Google Sheets formulas

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
  const maxColumns = clientReportSheet.getMaxColumns();
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

  // Option A
  // Set the width of the current column to 200 pixels.
  //clientReportSheet.setColumnWidth(c, 200);
  
  // Option B
  // Automatically resize all columns based on their content.
  //clientReportSheet.autoResizeColumns(1, maxColumns);

  // -------------------------------------------------------------
  // CARGA DE CONFIGURACIÓN CON FILTRO DE SERVICIOS
  // -------------------------------------------------------------
  const mappings = loadServiceMappings(trafiguraFlag, excelFlag);
  const maxRows = clientReportSheet.getMaxRows();

  let currentCol = 1;

  mappings.forEach(mapping => {
    // Encabezado Fila 1
    clientReportSheet.getRange(1, currentCol)
      .setValue(mapping.columnName)
      .setFontSize(10)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    // Fila 2 vacía o con fórmula
    clientReportSheet.getRange(2, currentCol).setValue("");
    if (mapping.formula) {
      clientReportSheet.getRange(2, currentCol).setFormula(mapping.formula);
    }

    // Formato de fechas o números en filas de datos
    if (maxRows > 2) {
      const dataRange = clientReportSheet.getRange(3, currentCol, maxRows - 2, 1);
      const dateType = mapping.dateType.toLowerCase();

      if (dateType === "date") {
        dataRange.setNumberFormat("yyyy-mm-dd");
      } else if (dateType === "number") {
        dataRange.setNumberFormat("0.00");
      }
    }

    // Colores de columna
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

  // After all columns have been created
  clientReportSheet.autoResizeColumns(1, currentCol - 1);

  // Add a small visual margin
  for (let c = 1; c < currentCol; c++) {
    clientReportSheet.setColumnWidth(
      c,
      clientReportSheet.getColumnWidth(c) + 60
    );
  }

  // Filtro automático
  const lastColUsed = clientReportSheet.getLastColumn();
  const lastRowUsed = clientReportSheet.getLastRow();

  if (lastColUsed > 0 && lastRowUsed >= 2) {
    const existingFilter = clientReportSheet.getFilter();
    if (existingFilter) {
      existingFilter.remove();
    }
    clientReportSheet.getRange(2, 1, lastRowUsed - 1, lastColUsed).createFilter();
  }

  // clientReportSheet.autoResizeColumns(1, maxColumns);
}

/**
 * Reads metadata configuration from Service_Config.
 * Filters columns against confirmed services from OperationSummary
 * and prevents duplicate column names.
 * 
 * @param {number} trafiguraFlag - 1 for Trafigura headers (Col B), 0 for standard headers (Col A).
 * @param {number} excelFlag - 1 for formula column M, 0 for formula column L.
 * @returns {Array<Object>} Mappings list.
 */
function loadServiceMappings(trafiguraFlag, excelFlag) {
  const configSS = SpreadsheetApp.openByUrl(
    //"https://docs.google.com/spreadsheets/d/1FZo4hMpiJOM-cfFbC10Wb5Nug27T77wF/edit#gid=2107217982"
    "https://docs.google.com/spreadsheets/d/1QqmHCH6OgZaV3kDXz8eRPpyhOffRlgIzH_nDlQ2urRI/edit#gid=2118177262"
  );
  
  const configSheet = configSS.getSheetByName(
  trafiguraFlag === 1
    ? 'Traf Master Tab (Traf Mine)'
    : 'Master Tab (All other Templates)'
  );

  if (!configSheet) {
    throw new Error("Sheet 'Master Tab (All other Templates) or Traf Master Tab (Traf Mine)' not found in the configuration spreadsheet.");
  }

  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) return [];

  const values = configSheet.getRange(2, 1, lastRow - 1, 28).getValues();
  // 2              Start row
  // 1              Start column 1 (A)
  // lastRow - 1    All rows from 2 to the last one
  // 13             End column to be used

  // 1. Obtener la lista de servicios confirmados desde OperationSummary.gs
  const confirmedServices = getConfirmedOperationsList();

  // 2. Set para evitar duplicados de nombres de columnas
  const seenColumnNames = new Set();

  // const targetColFormulaIndex = (trafiguraFlag === 1) ? 4 : 9;
  const targetFormulaIndex = (trafiguraFlag === 1) ? 4 : 9;

  const mappings = [];

  values.forEach(row => {
    let columnName = String(row[0] || "").trim();
    const serviceClass = String(row[1] || "").trim();
    if (serviceClass && serviceClass.toLowerCase() !== "general") {
      columnName = `${serviceClass}_${columnName}`;
  }

    // -------------------------------------------------------------
    // FILTRO 1: EVALUACIÓN DE SERVICIO (Columna F -> row[1])
    // -------------------------------------------------------------
    //const serviceClass = String(row[5] || "").trim();
    const serviceClassLower = serviceClass.toLowerCase();

    // Si tiene asignado un serviceClass específico (que no sea vacío ni "general")
    if (serviceClassLower !== "" && serviceClassLower !== "general") {
      // Admite celdas con servicios separados por comas (ej: "Sampling, Grading")
      const definedServices = serviceClassLower.split(',').map(s => s.trim());

      // Comprobar si coincide con alguno de los servicios confirmados en OperationSummary
      const isConfirmed = definedServices.some(svc => {
        for (let confirmed of confirmedServices) {
          if (confirmed.includes(svc) || svc.includes(confirmed)) {
            return true;
          }
        }
        return false;
      });

      // Si el servicio no está confirmado en las filas 68-99, omitir columna
      if (!isConfirmed) return;
    }

    // -------------------------------------------------------------
    // FILTRO 2: EVITAR NOMBRES VACÍOS O DUPLICADOS
    // -------------------------------------------------------------
    const colKey = columnName.toLowerCase();
    if (!columnName || seenColumnNames.has(colKey)) {
      return; // Salta si está vacío o si ya fue insertada una columna con ese nombre
    }
    seenColumnNames.add(colKey);

    // Registro de especificación
    mappings.push({
      columnName: columnName,
      validation: String(row[20] || "").trim(),
      serviceClass: serviceClass,
      dateType: String(row[18] || "").trim(),
      formula: String(row[targetFormulaIndex] || "").trim(),
      calcFlag: String(row[19] || "").trim()
    });
  });

  const serviceOrder = Array.from(confirmedServices).map(s => s.toLowerCase());

  mappings.sort((a, b) => {
    const idxA = serviceOrder.findIndex(s =>
      a.serviceClass.toLowerCase().includes(s)
    );

    const idxB = serviceOrder.findIndex(s =>
      b.serviceClass.toLowerCase().includes(s)
    );

    return idxA - idxB;
  });


  return mappings;
  
}


// 1. Obtener la lista de servicios confirmados desde OperationSummary.gs
const confirmedServices = getConfirmedOperationsList();

// LINEA A AGREGAR:
console.log("Servicios confirmados detectados:", Array.from(confirmedServices));