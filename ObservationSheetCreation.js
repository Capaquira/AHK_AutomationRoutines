function createObservationsSheet(new_google_sheet) {
  let obsSheet = new_google_sheet.insertSheet('Observations'); 

  // ===== MAIN HEADER =====
  obsSheet.getRange(2, 2, 1, 5)
    .merge()
    .setValue('OBSERVATIONS TAB')
    .setFontSize(24)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#fce8e6');

  // ===== SUB-HEADER =====
  obsSheet.getRange(3, 2, 1, 5)
    .merge()
    .setValue('UNITS TO LOOK INTO')
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#fce8e6');

  // ===== TABLE 1 =====
  obsSheet.getRange(5, 2, 1, 2)
    .merge()
    .setValue('Duplicate Check')
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#fce8e6');

  obsSheet.setRowHeight(5, 60); 
  obsSheet.getRange(5, 1, 1, obsSheet.getMaxColumns()).setWrap(true); 
  obsSheet.getRange(5, 1, 1, obsSheet.getMaxColumns()).setHorizontalAlignment('center'); 

  obsSheet.getRange(6, 2).setValue('Reception Duplicates').setBackground('#fce8e6');
  obsSheet.getRange(7, 2).setValue('=if(B7="No duplicates","0",counta(B7:B))');
  obsSheet.getRange(8, 2).setValue('=iferror(unique(filter(Client_Report_Reception!G:G,countif(Client_Report_Reception!G:G,Client_Report_Reception!G:G)>1)),"No Duplicates")');

  obsSheet.getRange(6, 3).setValue('Loading Duplicates').setBackground('#fce8e6');
  obsSheet.getRange(7, 3).setValue('=if(C7="No duplicates",0,counta(C7:C))');
  obsSheet.getRange(8, 3).setValue('=iferror(unique(filter(Client_Report_Loading!G:G,countif(Client_Report_Loading!G:G,Client_Report_Loading!G:G)>1)),"No Duplicates")');

  // ===== TABLE 2 =====
  obsSheet.getRange(5, 5).setValue('Units on the Ground - Received but NOT Loaded').setBackground('#fce8e6');
  obsSheet.getRange(5, 6).setValue('Mismatch Units - Units in Loading but NOT Reception').setBackground('#fce8e6');

  obsSheet.getRange(6, 5, 1, 2)
    .merge()
    .setValue('Unit ID')
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#fce8e6');

  obsSheet.getRange(7, 5).setValue('=iferror(if(E7="No duplicates","0",counta(E7:E)),"No units not loaded")');
  obsSheet.getRange(8, 5).setValue('=unique(filter(Client_Report_Loading!G3:G,countif(Client_Report_Loading!G3:G,Client_Report_Loading!G3:G)=0,Client_Report_Loading!G3:G<>""))');

  obsSheet.getRange(7, 6).setValue('=iferror(if(F7="No duplicates","0",counta(F7:F)),"No mismatch units")');
  obsSheet.getRange(8, 6).setValue('=unique(filter(Client_Report_Reception!G3:G,countif(Client_Report_Reception!G3:G,Client_Report_Reception!G3:G)=0,Client_Report_Reception!G3:G<>""))');

  // ===== GENERAL FORMATTING =====
  for (let col = 2; col <= 6; col++) {
    obsSheet.setColumnWidth(col, 200);
    obsSheet.getRange(5, col, obsSheet.getMaxRows()-4, 1)
      .setFontSize(12)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }

  return obsSheet;
}
