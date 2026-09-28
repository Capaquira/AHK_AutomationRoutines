  // =========================================================================
  // WEIGH METHOD SUMMARY
  // =========================================================================

  function getConfirmedWeighingMethods() {
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // 1. Batch read B65:E96 (32 rows x 4 columns)
    const startRow = 65;
    const numRows = 96 - 65 + 1; 
    const data = activeSheet.getRange(startRow, 2, numRows, 4).getValues();

    const weighingMethods = [];

    // 2. Iterate through rows and filter by "Weighing" and "Yes"
    for (let i = 0; i < data.length; i++) {
      const description = String(data[i][0] || '').trim(); // Column B
      const isConfirmed = String(data[i][3] || '').trim().toLowerCase(); // Column E

      if (description.toLowerCase().includes('weighing') && isConfirmed === 'yes') {
        let extractedText = description;
        
        // Extract text after the last "-"
        if (description.includes('-')) {
          const parts = description.split('-');
          extractedText = parts[parts.length - 1].trim();
        }

        // 3. AVOID DUPLICATES: Only push if the extracted text is non-empty and NOT already in the array
        if (extractedText && !weighingMethods.includes(extractedText)) {
          weighingMethods.push(extractedText);
        }
      }
    }

  // 4. Return grammatically joined summary string
  return formatNaturalList(weighingMethods);
  }