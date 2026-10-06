/// ===============================================================
/// Project: Template Generator – Custom Menu and Flag Management
/// ===============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Template Menu')
    .addItem('Google Sheet Template', 'tempGen')
    .addItem('Excel Template', 'excelTemGen')
    .addToUi();
}

// Standard template (Tem = 0)
function tempGen() {
  TemplateGenerator.CreateNewGoogleSheet(0);
}

// Excel template (Tem = 1)
function excelTemGen() {
  TemplateGenerator.CreateNewGoogleSheet(1);
}