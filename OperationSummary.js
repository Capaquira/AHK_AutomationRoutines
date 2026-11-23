function getConfirmedOperationsSummary() {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const startRow = 68;
  const numRows = 99 - 68 + 1;
  const data = activeSheet.getRange(startRow, 2, numRows, 5).getValues(); 

  const serviceMap = new Map(); // clave = Reception/Loading/Stand-alone, valor = lista de servicios

  for (let i = 0; i < data.length; i++) {
    const descriptionRaw = String(data[i][0] || '').trim();
    const isConfirmed = String(data[i][3] || '').trim().toLowerCase();
    const specification = String(data[i][4] || '').trim();

    if (isConfirmed !== 'yes') continue;
    if (descriptionRaw.startsWith('-')) continue; // ignorar subservicios

    // Determinar si es servicio común o especial
    let description;
    if (
      descriptionRaw.toLowerCase().startsWith('reception') ||
      descriptionRaw.toLowerCase().startsWith('loading') ||
      descriptionRaw.toLowerCase().startsWith('stock inspection')
    ) {
      description = descriptionRaw; // mantener completo
    } else {
      description = descriptionRaw.split(' ')[0]; // solo primera palabra
    }

    if (specification) {

        specification
          .split(',')
          .map(spec => spec.trim())
          .filter(Boolean)
          .forEach(spec => {

        const key = spec; // Reception, Loading, Stand-alone
        if (!serviceMap.has(key)) {
          serviceMap.set(key, new Set());
        }
        serviceMap.get(key).add(description);
      });
    } else {
      if (!serviceMap.has(description)) {
        serviceMap.set(description, new Set());
      }
    }
  }

  const results = [];
  for (const [location, services] of serviceMap.entries()) {
    if (location.toLowerCase() === 'stand-alone') {
      services.forEach(s => results.push(s));
    } else if (services.size > 0) {
      const serviceList = [...services].join(' / ');
      results.push(`${location} (${serviceList})`);
    } else {
      results.push(location);
    }
  }

  // Evitar duplicados de Reception/Loading simples si ya existen con detalle
  const hasReceptionDetail = results.some(r => r.startsWith('Reception ('));
  const hasLoadingDetail = results.some(r => r.startsWith('Loading ('));

  const filtered = results.filter(r => {
    if (r === 'Reception' && hasReceptionDetail) return false;
    if (r === 'Loading' && hasLoadingDetail) return false;
    return true;
  });

  return formatNaturalList(filtered);
}

function formatNaturalList(items) {

  if (!Array.isArray(items) || items.length === 0) {
    return "";
  }

  if (items.length === 1) return items[0];
  if (items.length === 2) return `${items[0]} and ${items[1]}`;
  
  const lastItem = items[items.length - 1];
  const initialItems = items.slice(0, -1).join(', ');
  return `${initialItems} and ${lastItem}`;
}


/**
 * Returns all confirmed operation identifiers used by
 * legacy filtering logic.
 *
 * Includes:
 * - Primary Services
 * - Operation descriptions
 * - Specifications
 *
 * @returns {Set<string>}
 */
function getConfirmedOperationsList() {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const startRow = 68;
  const numRows = 99 - 68 + 1;
  // Col B (índice 0), Col E confirmación (índice 3), Col F especificación (índice 4)
  const data = activeSheet.getRange(startRow, 2, numRows, 5).getValues(); 

  const confirmedKeys = new Set();

  for (let i = 0; i < data.length; i++) {
    const descriptionRaw = String(data[i][0] || '').trim();
    const isConfirmed = String(data[i][3] || '').trim().toLowerCase();
    const specification = String(data[i][4] || '').trim();

    if (isConfirmed !== 'yes') continue;
    if (descriptionRaw.startsWith('-')) continue; // Ignorar subservicios

    let description;
    const descLower = descriptionRaw.toLowerCase();

    if (
      descLower.startsWith('reception') ||
      descLower.startsWith('loading') ||
      descLower.startsWith('stock inspection')
    ) {
      description = descriptionRaw;
    } else {
      description = descriptionRaw.split(' ')[0];
    }

    // Registrar descripciones confirmadas en minúsculas para comparaciones uniformes
    confirmedKeys.add(descLower);
    confirmedKeys.add(description.toLowerCase());

    // Si tiene especificación (ej. Stand-alone, Reception, etc.) registrarla también
    if (specification) {

      specification
        .split(',')
        .map(spec => spec.trim())
        .filter(Boolean)
        .forEach(spec => {

        confirmedKeys.add(spec.toLowerCase());
      });
    }
  }

  return confirmedKeys;
}


/*** Trird Part of the code ***/


function getConfirmedPrimaryServices() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getActiveSheet();

  const primaryServices = [];

  const reception =
    String(sheet.getRange("E24").getValue())
      .trim()
      .toLowerCase();

  const intermediate =
    String(sheet.getRange("E28").getValue())
      .trim()
      .toLowerCase();

  const loading =
    String(sheet.getRange("E31").getValue())
      .trim()
      .toLowerCase();

  if (reception === "yes") {
    primaryServices.push("Reception");
  }

  if (intermediate === "yes") {
    primaryServices.push("Intermediate");
  }

  if (loading === "yes") {
    primaryServices.push("Loading");
  }

  return primaryServices;
}

/*** Forth Part */
// -----------------------------------------------------------------------------
// Secondary Service Normalization
// Converts Operation Summary descriptions into Service_Config values
// -----------------------------------------------------------------------------
function normalizeSecondaryService(descriptionRaw) {

  const text = String(descriptionRaw || "")
    .trim()
    .toLowerCase();

  if (text.includes("weighing - platform scale")) {
    return "Weighing - Platform Scale";
  }

  if (text.includes("weighing - weighbridge")) {
    return "Weighing - Weighbridge";
  }

  if (text.includes("3rd party")) {
    return "Advised Weights";
  }

  if (text.includes("mine weights")) {
    return "Mine weights";
  }

  if (text.includes("visual inspection")) {
    return "Visual Inspection";
  }

  if (text.includes("transport unit")) {
    return "Transport unit";
  }

  if (text.includes("unit seals")) {
    return "Unit seals";
  }

  if (text.includes("inspected vessel loading")) {
    return "Vessel loading";
  }

  if (text.includes("advised vessel info")) {
    return "Advised vessel info";
  }

  if (text.startsWith("sampling")) {
    return "Sampling";
  }

  if (text.startsWith("fractions portions")) {
    return "Fractions portions";
  }

  if (text.startsWith("grading")) {
    return "Grading";
  }

  if (text.startsWith("radiation")) {
    return "Radiation";
  }

  if (text.startsWith("reweigh")) {
    return "Reweighing";
  }

  if (text.startsWith("bagging")) {
    return "Bagging";
  }

  if (text.startsWith("de-bagging")) {
    return "Debagging";
  }

  if (text.startsWith("blending")) {
    return "Blending";
  }

  if (text.startsWith("stock inspection")) {
    return "Stock inspection";
  }

  return null;

}

// Secondary Service
/**
 * Returns confirmed Secondary Services grouped by
 * Primary Service.
 *
 * Example:
 * {
 *   Reception: Set([...]),
 *   Intermediate: Set([...]),
 *   Loading: Set([...])
 * }
 */
function getConfirmedPrimarySecondaryServices() {

  const activeSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getActiveSheet();

  const data =
    activeSheet
      .getRange(
        68,
        2,
        32,
        5
      )
      .getValues();

  const services = {
    Reception: new Set(),
    Intermediate: new Set(),
    Loading: new Set()
  };

  let currentPrimaryService = null;
  let currentGroupConfirmed = false;

  data.forEach(row => {

    const descriptionRaw =
      String(row[0] || "")
        .trim();

    const confirmation =
      String(row[3] || "")
        .trim()
        .toLowerCase();

    const descriptionLower =
      descriptionRaw.toLowerCase();

    // ----------------------------------------------------"stock inspection / holding certificate" ||
    // Detect Primary Service sections
    // ----------------------------------------------------

    // Assumes services remain grouped under their
    // corresponding Primary Service section.

    if (descriptionLower === "reception") {

      currentPrimaryService = "Reception";
      currentGroupConfirmed = (confirmation === "yes");

      return;

    }


    else if (descriptionLower === "loading") {

      currentPrimaryService = "Loading";
      currentGroupConfirmed = (confirmation === "yes");

      return;

    }

    // Current Intermediate services act as section identifiers
    // because there is no explicit "Intermediate" row
    // in Operation Summary.

    else if (
      descriptionLower ===
        "stock inspection / holding certificate" ||
      descriptionLower ===
        "sampling" ||
      descriptionLower.startsWith("radiation") ||
      descriptionLower.startsWith("bagging") ||
      descriptionLower.startsWith("de-bagging")
    ) {

      currentPrimaryService = "Intermediate";
      currentGroupConfirmed = (confirmation === "yes");

    }

    // Confirmation after detecyion
    if (confirmation !== "yes") {
      return;
    }


    if (!currentGroupConfirmed) {
      return;
    }
    // ----------------------------------------------------
    // Normal Secondary Services
    // ----------------------------------------------------

    const normalized =
      normalizeSecondaryService(
        descriptionRaw
      );

    if (
      normalized &&
      currentPrimaryService
    ) {
      services[
        currentPrimaryService
      ].add(normalized);
    }

  });

  return services;

}

