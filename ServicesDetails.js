/**
 * Apply summary formatting and data population
 * @param {GoogleAppsScript.Spreadsheet.Sheet} summarySheet - Target sheet to format
 * @param {GoogleAppsScript.Spreadsheet.Sheet} activeSheet - Source sheet for values
 */
function formatSummarySheet(summarySheet, activeSheet) {
  // Global sheet settings
  summarySheet.setHiddenGridlines(true);
  summarySheet.getRange('1:1000').setFontFamily('Montserrat');

  // Logo and header
  summarySheet.getRange('A1:I1').setBackground('#616161');
  summarySheet.getRange('H1').setBackground('#e83f4b');
  summarySheet.setRowHeight(1, 80);
  summarySheet.setColumnWidth(1, 30);

  summarySheet.getRange('A1')
    .setValue('ALFRED H KNIGHT')
    .setFontSize(14)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left');

  // Column B labels
  const bLabels = [
    ['Quality'], ['Commodity'], ['Receiver'],
    ['Site'], ['City'], ['Country'], ['Operations']
  ];
  summarySheet.getRange('B4:B10').setValues(bLabels).setFontSize(9).setFontWeight('bold');

  // Background colors
  const colors_redWhite = Array.from({ length: 7 }, (_, i) => [(i % 2 === 0) ? '#fce8e6' : '#ffffff']);
  summarySheet.getRange('B4:C10').setBackgrounds(colors_redWhite.map(c => [c[0], c[0]]));
  summarySheet.getRange('F4:G10').setBackgrounds(colors_redWhite.map(c => [c[0], c[0]]));

  // Column F labels
  const fLabels = [
    ['Packaging Type'], ['Weighing Method'], ['Tare Weight [MT]'],
    ['Supervisor'], ['Operations manager'], ['Sampling (%)'], ['Transport Type']
  ];
  summarySheet.getRange('F4:F10').setValues(fLabels).setFontSize(9).setFontWeight('bold');

  // Gray/white alternating backgrounds
  const colors_grayWhite = Array.from({ length: 7 }, (_, i) => [(i % 2 === 0) ? '#e0e0e0' : '#ffffff']);
  summarySheet.getRange('D4:E10').setBackgrounds(colors_grayWhite.map(c => [c[0], c[0]]));
  summarySheet.getRange('H4:I10').setBackgrounds(colors_grayWhite.map(c => [c[0], c[0]]));

  // Column D values
  const dSourceCells = ['E29', 'E28', 'E23', 'E25', 'E26', 'E27'];
  const dValues = dSourceCells.map(cellRef => [capitalizeWords(activeSheet.getRange(cellRef).getValue())]);
  summarySheet.getRange('D4:D9').setValues(dValues).setFontSize(9).setFontWeight('normal');

  const operationsSummary = getConfirmedOperationsSummary();
  summarySheet.getRange('D10').setValue(operationsSummary).setFontSize(9).setFontWeight('normal');

  // Column H rich text values
  const companyDomain = "ahkgroup.com";
  const weighingSummary = getConfirmedWeighingMethods();

  const rawE30 = capitalizeWords(String(activeSheet.getRange('E30').getValue()) || '').trim();
  const rawE31 = capitalizeWords(String(activeSheet.getRange('E31').getValue()) || '').trim();
  const rawE32 = capitalizeWords(String(activeSheet.getRange('E32').getValue()) || '').trim();
  const rawE33 = capitalizeWords(String(activeSheet.getRange('E33').getValue()) || '').trim();
  const rawE47 = String(activeSheet.getRange('E47').getValue() || '');
  const rawE49 = String(activeSheet.getRange('E49').getValue() || '');

  const supervisorList = extractIndividualNames(rawE47).slice(0, 2);
  const managerList = extractIndividualNames(rawE49).slice(0, 1);

  // Combine packaging type
  let combinedPackagingType = 'N/A';
  if (rawE30 && rawE30.toUpperCase() !== 'N/A') combinedPackagingType = `${rawE30} (Reception)`;
  if (rawE32 && rawE32.toUpperCase() !== 'N/A') combinedPackagingType = combinedPackagingType === 'N/A' ? rawE32 : `${combinedPackagingType} / ${rawE32} (Loading)`;

  // Combine transport type
  let combinedTransportType = 'N/A';
  if (rawE31 && rawE31.toUpperCase() !== 'N/A') combinedTransportType = `${rawE31} (Reception)`;
  if (rawE33 && rawE33.toUpperCase() !== 'N/A') combinedTransportType = combinedTransportType === 'N/A' ? rawE33 : `${combinedTransportType} / ${rawE33} (Loading)`;

  const hSourceMapping = [
    combinedPackagingType,
    weighingSummary,
    "",
    supervisorList,
    managerList,
    '',
    combinedTransportType
  ];

  const hRichTextValues = hSourceMapping.map((item, index) => {
    if ((index === 3 || index === 4) && Array.isArray(item) && item.length > 0) {
      const displayText = item.join(', ');
      const builder = SpreadsheetApp.newRichTextValue().setText(displayText);
      let searchStart = 0;
      item.forEach(name => {
        const nameStart = displayText.indexOf(name, searchStart);
        if (nameStart !== -1) {
          const nameEnd = nameStart + name.length;
          const email = extractOrInferEmail(name, companyDomain);
          if (email) builder.setLinkUrl(nameStart, nameEnd, `mailto:${email}`);
          searchStart = nameEnd;
        }
      });
      return [builder.build()];
    }
    return [SpreadsheetApp.newRichTextValue().setText(String(item || '')).build()];
  });

  summarySheet.getRange('H4:H10')
    .setRichTextValues(hRichTextValues)
    .setFontSize(9)
    .setFontWeight('normal')
    .setHorizontalAlignment('left');

  
  summarySheet.getRange('H6:I9')
    .setBackground('#FFFF00');

}





