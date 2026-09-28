/**
 * Creates a new Google Sheet named from active cell B2, immediately moves it
 * to a target Google Drive folder/Shared Drive, populates initial structure,
 * and seamlessly redirects the user via an HTML modal dialog.
 * 
 * Workflow:
 * 1. Reads dynamic file name from current sheet (Cell B2).
 * 2. Instantiates a new Spreadsheet and moves it to the target directory.
 * 3. Populates, formats, and structures all necessary tabs.
 * 4. Dynamically sorts tabs (Fixed Start, Dynamic Middle, Fixed End).
 * 5. Triggers a client-side JavaScript redirect via UI Modal.
 * 
 * @requires DESTINATION_FOLDER_ID - Valid Drive Folder ID with edit permissions.
 */

function CreateNewGoogleSheet(trafiguraFlag) {

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
  const DESTINATION_FOLDER_ID = "1abEPExDsu_w4-VPLy5JVckZf59m42QvS";

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
  const copiedActiveSheet = current_google_sheet.copyTo(new_google_sheet);

  // Set the copied tab's name to match the original sheet's name
  copiedActiveSheet.setName(current_google_sheet.getName());

  // Get current spreadsheet's URL and place it in cell I1 of the copied sheet
  const originalUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  copiedActiveSheet.getRange('I2')
    .setFormula(`=HYPERLINK("${originalUrl}", "Link to Original Sheet")`)
    .setFontColor('#1155cc')
    .setFontSize(15)
    .setFontWeight('bold');

  // Auto fit for column visibility
  // 9 represents Column I (A=1, B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9)
  copiedActiveSheet.setColumnWidth(9, 250);

  // =========================================================================
  // 4. REMOVE ALL BUTTONS & DRAWINGS FROM COPIED SHEET
  // =========================================================================

  // Retrieve array of all drawing objects (buttons, diagrams, drawings) in the copied tab
  const drawings = copiedActiveSheet.getDrawings();

  // Iterate over each drawing object and remove it from the sheet
  drawings.forEach(drawing => drawing.remove());

  // =========================================================================
  // Sheet02. COPY VALIDATION SHEET
  // =========================================================================

  const SOURCE_SPREADSHEET_ID = "11pgBEOENoeB8g0RxlNImbV5sGSL3v58S4LeRv-Mo3G0";
  const sourceSpreadsheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  const sourceSheet = sourceSpreadsheet.getSheetByName("Validation");

  let validationSheet = sourceSheet.copyTo(new_google_sheet);
  validationSheet.setName("Validation");

  // =========================================================================
  // Sheet03. COPY MERGE LOG SHEET
  // =========================================================================

  const SOURCE_SPREADSHEET_ID2 = "1qktFxlaQvxHdiGoJrWKqVUB6orxGwWd4jfX6Mkq-8FM";
  const sourceSpreadsheet2 = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID2);
  const sourceSheet2 = sourceSpreadsheet2.getSheetByName("Merge Log");

  let MergeLogSheet = sourceSheet2.copyTo(new_google_sheet);
  MergeLogSheet.setName("Merge Log");

  // =========================================================================
  // Sheet04. CREATE OBSERVATION SHEET
  // =========================================================================
  createObservationsSheet(new_google_sheet);

  // =========================================================================
  // Sheet05. CREATE SUMMARY SHEETS
  // =========================================================================

  // Then loop through confirmed operations and create extra sheets
  createOperationSheets(new_google_sheet);

  // =========================================================================
  // Sheet06. CREATE CLIENT REPORT SHEET
  // =========================================================================

  // Creation of main input file
  // To ensure default is 0 if not passed
  var flag = (trafiguraFlag === 1)? 1:0;
  createClientReportColumns(new_google_sheet, flag);

  // =========================================================================
  // Sheet07. CREATE "Client_Summary" SHEET
  // =========================================================================

  // Always create Client_Summary first
  let clientSheet = new_google_sheet.insertSheet('Client_Summary');
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  formatSummarySheet(clientSheet, activeSheet);

  clientSheet.getRange('B2')
    .setValue('Inspection Summary')
    .setFontSize(22)
    .setFontWeight('bold');

  // =========================================================================
  // Sheet08. COPY EXPORT SHEET
  // =========================================================================

  const SOURCE_SPREADSHEET_ID3 = "1THcO95rN0tTvGLuVWvNP2WfinrE3Izz2R3L5U1a5H0A";
  const sourceSpreadsheet3 = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID3);
  const sourceSheet3 = sourceSpreadsheet3.getSheetByName("Export");

  let ExportSheet = sourceSheet3.copyTo(new_google_sheet);
  ExportSheet.setName("Export");

  // =========================================================================
  // 5. CLEANUP DEFAULT SHEET
  // =========================================================================
  
  // Delete default Sheet1 if there are other sheets available
  const defaultTab = new_google_sheet.getSheets()[0];
  if (new_google_sheet.getSheets().length > 1) {
    new_google_sheet.deleteSheet(defaultTab);
  }

  // =========================================================================
  // 6. DYNAMIC TAB SORTING (START, MIDDLE, END)
  // =========================================================================
  
  const startSheets = ["Export", "Client_Summary", "Client_Report"];
  const endSheets = ["Observations", "Merge Log", "Validation", copiedActiveSheet.getName()];
  
  let currentIndex = 1;

  // Step A: Position the fixed start sheets in order
  startSheets.forEach(sheetName => {
    let sheet = new_google_sheet.getSheetByName(sheetName);
    if (sheet) {
      new_google_sheet.setActiveSheet(sheet);
      new_google_sheet.moveActiveSheet(currentIndex);
      currentIndex++;
    }
  });

  // Step B: Automatically position any dynamically generated middle sheets
  let allSheets = new_google_sheet.getSheets();
  let middleSheets = allSheets.filter(sheet => {
    let name = sheet.getName();
    return !startSheets.includes(name) && !endSheets.includes(name);
  });

  middleSheets.forEach(sheet => {
    new_google_sheet.setActiveSheet(sheet);
    new_google_sheet.moveActiveSheet(currentIndex);
    currentIndex++;
  });

  // Step C: Position the fixed end sheets in exact order
  endSheets.forEach(sheetName => {
    let sheet = new_google_sheet.getSheetByName(sheetName);
    if (sheet) {
      new_google_sheet.setActiveSheet(sheet);
      new_google_sheet.moveActiveSheet(currentIndex);
      currentIndex++;
    }
  });

  // Option: Flush changes to force immediate sync before opening URL
  SpreadsheetApp.flush();

  // Logging execution details
  Logger.log("File Name: " + new_template_name);
  Logger.log("Destination Folder: " + targetFolder.getName());
  Logger.log("File URL: " + new_google_sheet.getUrl());

  // =========================================================================
  // 7. USER INTERFACE & AUTOMATIC REDIRECTION
  // =========================================================================

  // SpreadsheetApp.getUi(): Retrieves the User Interface (UI) environment 
  // of the active spreadsheet, enabling alerts, custom menus, and dialog windows.
  // REF.: https://developers.google.com/apps-script/reference/base/ui
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