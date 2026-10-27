
  // Definición de la función
  function logGeneratedTemplate(templateName, templateSpreadsheet, flag_EX) {
    const logSS = SpreadsheetApp.openByUrl(
      "https://docs.google.com/spreadsheets/d/1Y8AVWTmiQDVUimgT4_OkpD6k3rT8Zm4AgyTVk6hGBu0/edit#gid=0"
    );
    const logSheet = logSS.getSheetByName("Log") || logSS.insertSheet("Log");




    // == Date and Time ==
    // LATAM  (GMT-4)
    // UK (GMT+1)

    // Exceuter Time
    const executerTimeZone = Intl.DateTimeFormat().resolvedOptions().timeZone; // under testing

    const dateLATAM = Utilities.formatDate(new Date(), "America/La_Paz", "yyyy-MM-dd HH:mm:ss");
    const dateUK = Utilities.formatDate(new Date(), "Europe/London", "yyyy-MM-dd HH:mm:ss");
    const dateExecuter = Utilities.formatDate(new Date(),executerTimeZone, "yyyy-MM-dd HH:mm:ss");
    console.log(Intl.DateTimeFormat().resolvedOptions());

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
  
  // Executer location
  //const spreadsheetTZ = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const spreadsheetTZ = Intl.DateTimeFormat().resolvedOptions().timeZone; // under testing

    // Agregar fila respetando el orden de columnas:
    // DateLATAM | DateUK | File Name | Version | User Request | File Link | Script User Executer
    logSheet.appendRow([dateLATAM, dateUK,dateExecuter, fileName, version, userRequest, fileType, fileUrl, scriptUserExecuter, spreadsheetTZ]);
  }