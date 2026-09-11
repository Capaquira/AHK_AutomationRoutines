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
    //To ensure default is 0 if not passed
    var flag = (trafiguraFlag === 1)? 1:0;
  createClientReportColumns(new_google_sheet, flag);

  // Then loop through confirmed operations and create extra sheets
  createOperationSheets(new_google_sheet);

  // =========================================================================
  // 6. CREATE AND CONFIGURE VALIDATION SHEET (DYNAMIC VALIDATION DATA)
  // =========================================================================

  const SOURCE_SPREADSHEET_ID = "11pgBEOENoeB8g0RxlNImbV5sGSL3v58S4LeRv-Mo3G0";
  const sourceSpreadsheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  const sourceSheet = sourceSpreadsheet.getSheetByName("Validation");

  let validationSheet = sourceSheet.copyTo(new_google_sheet);
  validationSheet.setName("Validation");

  // Move the copied sheet before the last sheet
  let totalSheets = new_google_sheet.getSheets().length;
  new_google_sheet.setActiveSheet(validationSheet);
  new_google_sheet.moveActiveSheet(totalSheets - 1);
 
  /************************************
   * MAIN OBS SHEET: OBSERVATIONS     *
   ************************************/

  let obsSheet = new_google_sheet.insertSheet('Observations'); 
  //REF: https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#insertsheetname

  // ===== MAIN HEADER =====
  // getRange(2, 2, 1, 5) → row 2, col 2 (B2), 1 row × 6 cols (B–F)
  obsSheet.getRange(2, 2, 1, 5) //REF: https://developers.google.com/apps-script/reference/spreadsheet/sheet#getrangerow,-column,-numrows,-numcolumns
    .merge() //REF: https://developers.google.com/apps-script/reference/spreadsheet/range#merge
    .setValue('OBSERVATIONS TAB') //REF: https://developers.google.com/apps-script/reference/spreadsheet/range#setvaluevalue
    .setFontSize(24) //REF: https://developers.google.com/apps-script/reference/spreadsheet/range#setfontsizefontsize
    .setFontWeight('bold') //REF: https://developers.google.com/apps-script/reference/spreadsheet/range#setfontweightweight
    .setHorizontalAlignment('center') //REF: https://developers.google.com/apps-script/reference/spreadsheet/range#sethorizontalalignmentalignment
    .setBackground('#fce8e6'); //REF: https://developers.google.com/apps-script/reference/spreadsheet/range#setbackgroundcolor

  // ===== SUB-HEADER =====
  // getRange(3, 2, 1, 5) → row 3, col 2 (B3), 1 row × 6 cols (B–F)
  obsSheet.getRange(3, 2, 1, 5)
    .merge()
    .setValue('UNITS TO LOOK INTO')
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#fce8e6');

  // ===== TABLE 1 =====
  // Title → Duplicate Check
  // getRange(5, 2, 1, 2) → row 5, col 2 (B5), 1 row × 2 cols (B–C)
  obsSheet.getRange(5, 2, 1, 2)
    .merge()
    .setValue('Duplicate Check')
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#fce8e6');

  // Set row height for row 5
  obsSheet.setRowHeight(5, 60); 
  //REF: https://developers.google.com/apps-script/reference/spreadsheet/sheet#setrowheightrow,-height

  // Apply wrap text to all cells in row 5
  obsSheet.getRange(5, 1, 1, obsSheet.getMaxColumns()).setWrap(true); 
  //REF: https://developers.google.com/apps-script/reference/spreadsheet/range#setwrapwrap

  // Apply horizontal center to all cells in row 5
  obsSheet.getRange(5, 1, 1, obsSheet.getMaxColumns()).setHorizontalAlignment('center'); 
  //REF: https://developers.google.com/apps-script/reference/spreadsheet/range#sethorizontalalignmentalignment

  // Subtitle → Reception Duplicates
  obsSheet.getRange(6, 2).setValue('Reception Duplicates').setBackground('#fce8e6');
  obsSheet.getRange(7, 2).setValue('=if(B7="No duplicates","0",counta(B7:B))');
  obsSheet.getRange(8, 2).setValue('=iferror(unique(filter(Client_Report_Reception!G:G,countif(Client_Report_Reception!G:G,Client_Report_Reception!G:G)>1)),"No Duplicates")');

  // Subtitle → Loading Duplicates
  obsSheet.getRange(6, 3).setValue('Loading Duplicates').setBackground('#fce8e6');
  obsSheet.getRange(7, 3).setValue('=if(C7="No duplicates",0,counta(C7:C))');
  obsSheet.getRange(8, 3).setValue('=iferror(unique(filter(Client_Report_Loading!G:G,countif(Client_Report_Loading!G:G,Client_Report_Loading!G:G)>1)),"No Duplicates")');

  // ===== TABLE 2 =====
  // Title → Units on the Ground - Received but NOT Loaded
  obsSheet.getRange(5, 5).setValue('Units on the Ground - Received but NOT Loaded').setBackground('#fce8e6');
  // Title → Mismatch Units - Units in Loading but NOT Reception
  obsSheet.getRange(5, 6).setValue('Mismatch Units - Units in Loading but NOT Reception').setBackground('#fce8e6');

  // Subtitle → Unit ID
  // getRange(6, 5, 1, 2) → row 6, col 5 (E6), 1 row × 2 cols (E–F)
  obsSheet.getRange(6, 5, 1, 2)
    .merge()
    .setValue('Unit ID')
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#fce8e6');

  // Formulas → Units not loaded
  obsSheet.getRange(7, 5).setValue('=iferror(if(E7="No duplicates","0",counta(E7:E)),"No units not loaded")');
  obsSheet.getRange(8, 5).setValue('=unique(filter(Client_Report_Loading!G3:G,countif(Client_Report_Loading!G3:G,Client_Report_Loading!G3:G)=0,Client_Report_Loading!G3:G<>""))');

  // Formulas → Mismatch units
  obsSheet.getRange(7, 6).setValue('=iferror(if(F7="No duplicates","0",counta(F7:F)),"No mismatch units")');
  obsSheet.getRange(8, 6).setValue('=unique(filter(Client_Report_Reception!G3:G,countif(Client_Report_Reception!G3:G,Client_Report_Reception!G3:G)=0,Client_Report_Reception!G3:G<>""))');

  // ===== GENERAL FORMATTING =====
  // Applies to columns B–F (includes row 5 titles and all rows below)
  for (let col = 2; col <= 6; col++) { //REF: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Statements/for
    // col = 2 → B, col = 3 → C, col = 4 → D, col = 5 → E, col = 6 → F
    obsSheet.setColumnWidth(col, 200); //REF: https://developers.google.com/apps-script/reference/spreadsheet/sheet#setcolumnwidthcolumn,-width
    obsSheet.getRange(5, col, obsSheet.getMaxRows()-4, 1)
      .setFontSize(12)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
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