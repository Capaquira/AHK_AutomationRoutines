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
function createClientReportColumns(new_google_sheet, trafiguraFlag) {
  let clientReportSheet = new_google_sheet.getSheetByName('Client_Report');
  if (!clientReportSheet) {
    clientReportSheet = new_google_sheet.insertSheet('Client_Report');
  }

  const flag = parseInt(trafiguraFlag || 0, 10);
  console.log("Generating Client_Report with TrafiguraFlag:", flag);

  // Formato de encabezado fila 1
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
  // CARGA DE CONFIGURACIÓN CON FILTRO DE SERVICIOS
  // -------------------------------------------------------------
  const mappings = loadServiceMappings(flag);
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
}

/**
 * Reads metadata configuration from Service_Config.
 * Filters columns against confirmed services from OperationSummary
 * and prevents duplicate column names.
 * 
 * @param {number} trafiguraFlag - 1 for Col B, 0 for Col A.
 * @returns {Array<Object>} Mappings list.
 */
function loadServiceMappings(trafiguraFlag) {
  const configSS = SpreadsheetApp.openByUrl(
    "https://docs.google.com/spreadsheets/d/1FZo4hMpiJOM-cfFbC10Wb5Nug27T77wF/edit#gid=2107217982"
  );
  const configSheet = configSS.getSheetByName('Service_Config');

  if (!configSheet) {
    throw new Error("Sheet 'Service_Config' not found in the configuration spreadsheet.");
  }

  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) return [];

  const values = configSheet.getRange(2, 1, lastRow - 1, 12).getValues();

  // 1. Obtener la lista de servicios confirmados desde OperationSummary.gs
  const confirmedServices = getConfirmedOperationsList();

  // 2. Set para evitar duplicados de nombres de columnas
  const seenColumnNames = new Set();

  const targetColIndex = (trafiguraFlag === 1) ? 1 : 0;
  const mappings = [];

  values.forEach(row => {
    // Definir nombre de columna (Col B para Trafigura con respaldo en Col A, o Col A para estándar)
    let columnName = String(row[targetColIndex] || "").trim();
    if (trafiguraFlag === 1 && !columnName) {
      columnName = String(row[0] || "").trim();
    }

    // -------------------------------------------------------------
    // FILTRO 1: EVALUACIÓN DE SERVICIO (Columna F -> row[5])
    // -------------------------------------------------------------
    const serviceClass = String(row[5] || "").trim();
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

      // Si el servicio no está confirmado en las filas 64-95, omitir columna
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
      validation: String(row[9] || "").trim(),
      serviceClass: serviceClass,
      dateType: String(row[7] || "").trim(),
      formula: String(row[11] || "").trim(),
      calcFlag: String(row[8] || "").trim()
    });
  });

  return mappings;
}


// 1. Obtener la lista de servicios confirmados desde OperationSummary.gs
  const confirmedServices = getConfirmedOperationsList();

  // LINEA A AGREGAR:
  console.log("Servicios confirmados detectados:", Array.from(confirmedServices));