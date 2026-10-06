
  // Definición de la función
  function logGeneratedTemplate(templateName, templateSpreadsheet, flag_EX) {
    const logSS = SpreadsheetApp.openByUrl(
      "https://docs.google.com/spreadsheets/d/1Y8AVWTmiQDVUimgT4_OkpD6k3rT8Zm4AgyTVk6hGBu0/edit#gid=0"
    );
    const logSheet = logSS.getSheetByName("Log") || logSS.insertSheet("Log");

    // Fecha y hora
    const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss");

    // Usuario que ejecuta el script
    const scriptUserExecuter = Session.getActiveUser().getEmail();

    // Nombre del archivo
    const fileName = templateName;

    // User Request: si el nombre contiene "TRAFIGURA"
    const userRequest = fileName.toUpperCase().includes("TRAFIGURA") ? "TRAFIGURA" : "COMMON";
    const fileType = flag_EX === 1 ? "EXCEL" : "GOOGLE SHEET";

    // URL del archivo
    const fileUrl = templateSpreadsheet.getUrl();

    // Versión: buscar último registro en la columna 4 (Version)
    let version = "V_0000000001";
    const lastRow = logSheet.getLastRow();
    if (lastRow > 1) { // la fila 1 son encabezados
      const lastVersion = logSheet.getRange(lastRow, 4).getValue();
      if (lastVersion) {
        // Elimina el prefijo "V_" y convierte a número
        const numericPart = parseInt(lastVersion.replace("V_", ""), 10);
        const nextNum = numericPart + 1;
        version = "V_" + Utilities.formatString("%010d", nextNum);
      }
    }

    // Agregar fila respetando el orden de columnas:
    // Date | Time | File Name | Version | User Request | File Link | Script User Executer
    logSheet.appendRow([date, time, fileName, version, userRequest, fileType, fileUrl, scriptUserExecuter]);
  }