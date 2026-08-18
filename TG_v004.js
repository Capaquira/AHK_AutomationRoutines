/**
 * Creates a new Google Sheet named from active cell B2, immediately moves it
 * to a target Google Drive folder/Shared Drive, populates initial structure,
 * and seamlessly redirects the user via an HTML modal dialog.
 * 
 * Workflow:
 * 1. Reads dynamic file name from current sheet (Cell B2).
 * 2. Instantiates a new Spreadsheet and moves it to the target directory.
 * 3. Formats default 'Export' tab with standard headers.
 * 4. Triggers a client-side JavaScript redirect via UI Modal.
 * 
 * @requires DESTINATION_FOLDER_ID - Valid Drive Folder ID with edit permissions.
 */

function CreateNewGoogleSheet() {

  // =========================================================================
  // 1. READ DATA FROM CURRENT SHEET
  // =========================================================================

  // SpreadsheetApp: The primary service used to interact with Google Sheets.
  // getActiveSpreadsheet(): Fetches the currently open spreadsheet workbook.
  // getActiveSheet(): Captures the specific active tab where the user is executing the script.
  // REF.: https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet
  const current_google_sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // getRange('B2'): Selects cell B2 within the active sheet.
  // getValue(): Retrieves the text or numerical value inside cell B2 to use as the new file name.
  const new_template_name = current_google_sheet.getRange('B2').getValue();

  // Set up destination
  const DESTINATION_FOLDER_ID = "1FvXBvi8yHsNWGTEuubZnCPsScFRklgXI";


  // =========================================================================
  // 2. CREATE AND MOVE FILE
  // =========================================================================
  
  // Step A: Create standalone sheet in root
  // SpreadsheetApp.create(name): Generates a brand-new standalone Google Sheet 
  // saved at the root directory of the user's Google Drive using the provided string name.
  // Returns the newly created 'Spreadsheet' object instance.
  // REF.: https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app#create(String)
  const new_google_sheet = SpreadsheetApp.create(new_template_name);

  // Step B: Move immediately to the target folder/Shared Drive
  // Bridge Spreadsheet object to DriveApp File reference using its unique ID
  // REF: https://developers.google.com/apps-script/reference/drive/drive-app#getfilebyidid
  const targetFile = DriveApp.getFileById(new_google_sheet.getId());
  // Fetch target folder (validates folder existence and user access)
  // REF: https://developers.google.com/apps-script/reference/drive/drive-app#getfolderbyidid
  const targetFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  // Relocate file pointer to destination (inherits permissions & Shared Drive policies)
  // REF: https://developers.google.com/apps-script/reference/drive/file#movetofolder
  targetFile.moveTo(targetFolder);
  
  // Logger.log(): Prints debugging output directly to the Apps Script execution logs.
  Logger.log("Assigned Template Name: " + new_template_name);
  Logger.log("Your new sheet is ready! Open it here: " + new_google_sheet.getUrl());

  // =========================================================================
  // 3. COPY SHEET CONTENT TO NEW FILE
  // =========================================================================

  // Copy active tab (data, formatting, formulas) into the new spreadsheet
  const copiedSheet = current_google_sheet.copyTo(new_google_sheet);

  // Set the copied tab's name to match the original sheet's name
  copiedSheet.setName(current_google_sheet.getName());

  // Get current spreadsheet's URL and place it in cell I1 of the copied sheet
  const originalUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  copiedSheet.getRange('I2')
    .setFormula(`=HYPERLINK("${originalUrl}", "Link to Original Sheet")`)
    .setFontColor('#1155cc')
    .setFontSize(15)
    .setFontWeight('bold');

  // Auto fit for column visibility
  // 9 represents Column I (A=1, B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9)
    copiedSheet.setColumnWidth(9, 250);

  //Main obs sheet "Observations"
  let obsSheet = new_google_sheet.insertSheet('Observations');

  // =========================================================================
  // 4. REMOVE ALL BUTTONS & DRAWINGS FROM COPIED SHEET
  // =========================================================================

  // Retrieve array of all drawing objects (buttons, diagrams, drawings) in the copied tab
  const drawings = copiedSheet.getDrawings();

  // Iterate over each drawing object and remove it from the sheet
  drawings.forEach(drawing => drawing.remove());

  // =========================================================================
  // 5. CREATE "Client_Summary" SHEET
  // =========================================================================

  // 1. Insert the new target sheet
  // REF: https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#insertsheetsheetname
  
  // Always create Client_Summary first
  let clientSheet = new_google_sheet.insertSheet('Client_Summary');
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  formatSummarySheet(clientSheet, activeSheet);

  clientSheet.getRange('B2')
    .setValue('Inspection Summary')
    .setFontSize(22)
    .setFontWeight('bold');

  //Creation of main input file
  createClientReportSheets(new_google_sheet)

  // Then loop through confirmed operations and create extra sheets
  createOperationSheets(new_google_sheet);

  // =========================================================================
  // 6. CREATE AND CONFIGURE VALIDATION SHEET (DYNAMIC VALIDATION DATA)
  // =========================================================================

  // 1. Insert the new "Validation" tab into the target spreadsheet
  // REF: https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#insertsheetsheetname
  const validationSheet = new_google_sheet.insertSheet('Validation');
  
  // Apply "Montserrat" to the entire grid (including empty cells)
  validationSheet.getRange('1:1000').setFontFamily('Montserrat')
  
  // Bands formating to target the range from A1 to F30 (or A1:F to cover all rows)
  const validationRange = validationSheet.getRange('A1:G30');

  validationRange.applyRowBanding()
    .setHeaderRowColor('#f38181')   // Header row background (Light Red)
    .setFirstRowColor('#ffffff')    // First alternating row (White)
    .setSecondRowColor('#fce8e6');  // Second alternating row (Light Pink / #f4c7c3)
  
  // Disables the default gridlines for the entire active tab
  validationSheet.setHiddenGridlines(true);

  // 2. Set header titles based on Validation.csv across columns A to F
  /*
   * getRange(row, column, numRows, numColumns) Parameters Breakdown:
   * - 1 (row): Starts at Row 1 (header row)
   * - 1 (column): Starts at Column 1 (Column A)
   * - 1 (numRows): Selects 1 row vertically
   * - 6 (numColumns): Selects 6 consecutive columns horizontally (A1:F1)
   * REF: https://developers.google.com/apps-script/reference/spreadsheet/sheet#getrangerow,-column,-numrows,-numcolumns
   */
  validationSheet.getRange(1, 1, 1, 6).setValues([[
    'Cargo condition', 
    'Vehicle types', 
    'Surveyor Symbol Colour', 
    'Surveyor Symbol Shape', 
    'Transporter', 
    'Destination'
  ]]).setFontWeight('bold').setFontColor('#ffffff').setHorizontalAlignment('center');

  // 3. Populate initial validation values for Column A (Cargo condition)
  // REF: https://developers.google.com/apps-script/reference/spreadsheet/range#setvaluesvalues
  validationSheet.getRange('A2:A24').setValues([
    ['Accident cargo'],
    ['Broken Strap'],
    ['Discoloration'],
    ['Missing seal bag'], 
    ['Missing seal tarp'],
    ['Oxidation'],
    ['Stolen Cargo'],
    ['Torn bag'], 
    ['Visually rejected'],
    ['Wet'],
    ['Open bag'],
    ['Over weight'], 
    ['Underweight'],
    ['Double bagged'],
    ['Re bagged'],
    ['Opened on top'], 
    ['Re-sealed'],
    ['Missing mine seal'],
    ['Underweight/torn'],
    ['Partially open'], 
    ['All in order'],
    ['Sweepings'],
    ['Replaced seal']
  ]);

  // 4. Populate values for Column B (Vehicle types)
  validationSheet.getRange('B2:B3').setValues([
    ['Flat bed'], 
    ['Container']
  ]);

  // 5. Populate values for Column C (Surveyor Symbol Colour)
  validationSheet.getRange('C2:C5').setValues([
    ['Black'], 
    ['Yellow'], 
    ['Red'], 
    ['Blue']
  ]);

  // 6. Populate values for Column D (Surveyor Symbol Shape)
  validationSheet.getRange('D2:D5').setValues([
    ['Star'], 
    ['Circle'], 
    ['Triangle'], 
    ['Square']
  ]);

  // 7. Populate default placeholder entries for Columns E & F (Transporter & Destination)
  // REF: https://developers.google.com/apps-script/reference/spreadsheet/range#setvaluevalue
  validationSheet.getRange('E2').setValue('<Insert data>');
  validationSheet.getRange('F2').setValue('<Insert data>');

  // 8. Auto-adjust column widths across all 6 validation lists for visibility
  // REF: https://developers.google.com/apps-script/reference/spreadsheet/sheet#autoresizecolumncolumnposition
  validationSheet.getRange(1, 1, 1, 8).setWrap(false); //Force single-line display (no line wrapping) across headers A1:F1
  for (let col = 1; col <= 6; col++) {
    validationSheet.setColumnWidth(col, 125);
  }

  // =========================================================================
  // 6. APPLY DATA & FORMATTING (IN DEFINITIVE LOCATION)
  // =========================================================================
  const exportTab = new_google_sheet.getSheets()[0];
  exportTab.setName('Export');
  exportTab.getRange(1, 1, 1, 4).setValues([['ID', 'Date', 'Customer Name', 'Total Amount']]);

  // Option: Flush changes to force immediate sync before opening URL
  SpreadsheetApp.flush();

  // Logging execution details
  Logger.log("File Name: " + new_template_name);
  Logger.log("Destination Folder: " + targetFolder.getName());
  Logger.log("File URL: " + new_google_sheet.getUrl());

  // =========================================================================
  // 6. USER INTERFACE & AUTOMATIC REDIRECTION
  // =========================================================================

  // SpreadsheetApp.getUi(): Retrieves the User Interface (UI) environment 
  // of the active spreadsheet, enabling alerts, custom menus, and dialog windows.
  //REF.: https://developers.google.com/apps-script/reference/base/ui
  const ui = SpreadsheetApp.getUi();
  
  // .getUrl(): A method of the Spreadsheet object that fetches the unique 
  // web address (URL) required to access the newly created file.
  const url = new_google_sheet.getUrl();
  
  // HtmlService.createHtmlOutput(html): Interprets and builds client-side HTML/JavaScript code.
  // - 'window.open(url, "_blank")': Browser JavaScript command that opens the URL in a new tab.
  // - 'google.script.host.close()': Apps Script client API that automatically closes the modal dialog.
  // - setWidth() / setHeight(): Sets the pop-up window dimensions in pixels.
  const htmlOutput = HtmlService
    .createHtmlOutput('<script>window.open("' + url + '", "_blank"); google.script.host.close();</script>')
    .setWidth(300)
    .setHeight(80);
    
  // ui.showModalDialog(userInterface, title): Renders the HTML pop-up window on screen,
  // temporarily focusing over the spreadsheet to execute the client-side redirection script.
  ui.showModalDialog(htmlOutput, 'Opening new template...');
}