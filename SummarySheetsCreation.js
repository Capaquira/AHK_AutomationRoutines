function createOperationSheets(new_google_sheet) {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // --- Helper: normalize transport type ---
  function normalizeTransportType(value) {
    if (!value) return null;
    const firstWord = String(value).trim().split(" ")[0].toLowerCase();
    if (firstWord === "n/a") return null;
    if (firstWord === "truck") return "Truck";
    if (firstWord === "train") return "Train";
    if (firstWord === "container") return "Container";
    return firstWord.charAt(0).toUpperCase() + firstWord.slice(1);
  }

  // --- Helper: create a tab ---
  function createTab(spreadsheet, sourceSheet, sheetName, title) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
    formatSummarySheet(sheet, sourceSheet);
    sheet.getRange("B2").setValue(title).setFontSize(22).setFontWeight("bold");
    Logger.log("Created tab: " + sheetName);
  }

  // --- Helper: evaluate a range with flexible matching ---
  function evaluateRange(spreadsheet, sourceSheet, startRow, endRow, mappings) {
    const numRows = endRow - startRow + 1;
    const descriptions = sourceSheet.getRange(startRow, 2, numRows, 1).getValues(); // Col B
    const confirmations = sourceSheet.getRange(startRow, 5, numRows, 1).getValues(); // Col E

    for (let i = 0; i < numRows; i++) {
      const description = String(descriptions[i][0] || "").trim().toLowerCase();
      const confirmed = String(confirmations[i][0] || "").trim().toLowerCase();

      if (confirmed === "yes") {
        for (let key in mappings) {
          if (description.startsWith(key.toLowerCase()) || description.includes(key.toLowerCase())) {
            const mapping = mappings[key];
            createTab(spreadsheet, sourceSheet, mapping.sheetName, mapping.title);
          }
        }
      }
    }
  }

  // --- Helper: check transport type with confirmation in rows 123–130 ---
  function checkTransportSummary(spreadsheet, sourceSheet, rowValue, label) {
    const transportValue = String(sourceSheet.getRange(rowValue, 5).getValue() || "").trim(); // E31 or E33
    if (!transportValue || transportValue.toLowerCase() === "n/a") return;

    const normalizedTransport = normalizeTransportType(transportValue);

    // Map transport type to expected confirmation phrase
    let expectedPhrase = "";
    if (normalizedTransport === "Truck") expectedPhrase = "truck x truck";
    else if (normalizedTransport === "Train") expectedPhrase = "train x train";
    else if (normalizedTransport === "Container") expectedPhrase = "container summary";

    // Scan rows 123–130 for confirmation
    const descriptions = sourceSheet.getRange(123, 2, 8, 1).getValues(); // Col B
    const confirmations = sourceSheet.getRange(123, 5, 8, 1).getValues(); // Col E

    let confirmed = false;
    for (let i = 0; i < descriptions.length; i++) {
      const desc = String(descriptions[i][0] || "").trim().toLowerCase();
      const conf = String(confirmations[i][0] || "").trim().toLowerCase();
      if (desc === expectedPhrase && conf === "yes") {
        confirmed = true;
        break;
      }
    }

    if (confirmed) {
      createTab(
        spreadsheet,
        sourceSheet,
        `${label}_${normalizedTransport}_Summary`,
        `${label} ${normalizedTransport} Summary`
      );
    }
  }

  // --- Range 1 & 3: Reception and Loading Transport Types ---
  checkTransportSummary(new_google_sheet, activeSheet, 31, "Reception");
  checkTransportSummary(new_google_sheet, activeSheet, 33, "Loading");

  // --- Range 2: Operations (Row 64–95) ---
  evaluateRange(new_google_sheet, activeSheet, 64, 95, {
    "reception": { sheetName: "Client_Reception_Summary", title: "Reception Summary" },
    "stock inspection": { sheetName: "Client_Stock_Inspection_Summary", title: "Stock Inspection Summary" },
    "loading": { sheetName: "Client_Loading_Summary", title: "Loading Summary" },
    "sampling": { sheetName: "Client_Sampling_Summary", title: "Sampling Summary" },
    "bagging": { sheetName: "Client_Bagging_Summary", title: "Bagging Summary" },
    "fractions portions": { sheetName: "Client_Fractions_Portions_Summary", title: "Fractions Portions Summary" },
    "grading": { sheetName: "Client_Grading_Summary", title: "Grading Summary" },
    "radiation": { sheetName: "Client_Radiation_Summary", title: "Radiation Summary" },
    "reweigh": { sheetName: "Client_Reweigh_Summary", title: "Reweigh Summary" },
    "de-bagging": { sheetName: "Client_Debagging_Summary", title: "De-bagging Summary" },
    "blending": { sheetName: "Client_Blending_Summary", title: "Blending Summary" }
  });

}
