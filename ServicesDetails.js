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
  
  // == Old Range ==
  // E23  Client 1
  // E25  Site
  // E26  Location
  // E27  Country
  // E28  Commodity
  // E29  Quality

  // == New Range ==
  // E16  AHK Template QC Code
  // E17  Client 1
  // E19  Site
  // E20  Location
  // E21  Country
  // E22  Commodity
  // E23  Quality

  var QCtemplateCode =activeSheet.getRange('E16') // To be used in the export tap

  // const dSourceCells = ['E29', 'E28', 'E23', 'E25', 'E26', 'E27'];
  const dSourceCells = ['E23', 'E22', 'E17', 'E19', 'E20', 'E21'];
  const dValues = dSourceCells.map(cellRef => [capitalizeWords(activeSheet.getRange(cellRef).getValue())]);
  summarySheet.getRange('D4:D9').setValues(dValues).setFontSize(9).setFontWeight('normal');

  const operationsSummary = getConfirmedOperationsSummary();
  summarySheet.getRange('D10').setValue(operationsSummary).setFontSize(9).setFontWeight('normal');

  // Column H rich text values
  const companyDomain = "ahkgroup.com";
  const weighingSummary = getConfirmedWeighingMethods();
  
  // == Old range ==

  // E30  Packaging Type (Reception)
  // E31  Transport Type (Reception)
  // E32  Packaging Type (Loading)
  // E33  Transport Type (Loading)

  // const rawE30 = capitalizeWords(String(activeSheet.getRange('E30').getValue()) || '').trim();
  // const rawE31 = capitalizeWords(String(activeSheet.getRange('E31').getValue()) || '').trim();
  // const rawE32 = capitalizeWords(String(activeSheet.getRange('E32').getValue()) || '').trim();
  // const rawE33 = capitalizeWords(String(activeSheet.getRange('E33').getValue()) || '').trim();

  // == New range ==
  
  // Reception
  // E25  Packaging Type (Reception)
  // E26  Transport Type (Reception)
  // E27  Reception - Cargo Transport Unit (CTU) .................. Where do we use this record?
  // ---> It is spected to be used for more columns selection like Cointainer_ID

  const rawE25 = capitalizeWords(String(activeSheet.getRange('E25').getValue()) || '').trim();
  const rawE26 = capitalizeWords(String(activeSheet.getRange('E26').getValue()) || '').trim();
  const rawE27 = capitalizeWords(String(activeSheet.getRange('E27').getValue()) || '').trim();
  
  // Intermediate
  // E29  Intermediate - Service .................. Where do we use this record?
  // E30  Intermediate - Packaging Type
  const rawE29 = capitalizeWords(String(activeSheet.getRange('E29').getValue()) || '').trim();
  const rawE30 = capitalizeWords(String(activeSheet.getRange('E30').getValue()) || '').trim();

  // Loading
  // E32  Vessel Loading .................. Where do we use this record?
  // E33  Packaging Type (Loading)
  // E34  Transport Type (Loading)

  // New Loading Range
  // E32  Loading - Packaging Type	
  // E33  Loading - Transport Type		
  // E34  Loading - Cargo Transport Unit (CTU)		
  // E35  Vessel Loading 		

  const rawE32 = capitalizeWords(String(activeSheet.getRange('E32').getValue()) || '').trim();
  const rawE33 = capitalizeWords(String(activeSheet.getRange('E33').getValue()) || '').trim();
  const rawE34 = capitalizeWords(String(activeSheet.getRange('E34').getValue()) || '').trim();
  const rawE35 = capitalizeWords(String(activeSheet.getRange('E35').getValue()) || '').trim();

  // Contac info
  const rawE51 = String(activeSheet.getRange('E51').getValue() || '');
  const rawE53 = String(activeSheet.getRange('E53').getValue() || '');
  const supervisorList = extractIndividualNames(rawE51).slice(0, 2);
  const managerList = extractIndividualNames(rawE53).slice(0, 1);

  // Combine packaging type
  const packagingTypes = [];
  if (rawE25 && rawE25.toUpperCase() !== 'N/A') {packagingTypes.push(`${rawE25} (Reception)`);}
  if (rawE30 && rawE30.toUpperCase() !== 'N/A') {packagingTypes.push(`${rawE30} (Intermediate)`);}
  if (rawE32 && rawE32.toUpperCase() !== 'N/A') {packagingTypes.push(`${rawE32} (Loading)`);}
  const combinedPackagingType = packagingTypes.length
  ? packagingTypes.join(' / ')
  : 'N/A';

  // Combine transport type
  const transportTypes = [];
  if (rawE26 && rawE26.toUpperCase() !== 'N/A') {transportTypes.push(`${rawE26} (Reception)`);}
  if (rawE33 && rawE33.toUpperCase() !== 'N/A') {transportTypes.push(`${rawE33} (Loading)`);}
  const combinedTransportType = transportTypes.length
  ? transportTypes.join(' / ')
  : 'N/A';

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





