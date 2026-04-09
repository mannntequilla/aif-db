function importLatestMyCaseLeadsReportFromDrive() {
  const folderId = '1DBC-0j9nnO20hReA_mjWQKoajPPaX9Zz';
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();

  let latestFile = null;
  let latestUpdated = 0;

  while (files.hasNext()) {
    const file = files.next();
    if (!file.getName().toLowerCase().endsWith('.csv')) continue;

    const updated = file.getLastUpdated().getTime();
    if (updated > latestUpdated) {
      latestUpdated = updated;
      latestFile = file;
    }
  }

  if (!latestFile) {
    throw new Error('No encontré ningún CSV en la carpeta.');
  }

  const csvText = latestFile.getBlob().getDataAsString();
  const rows = Utilities.parseCsv(csvText);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = CONFIG.sheets.rawMyCaseLeadsReport;
  if (!sheetName) {
    throw new Error('CONFIG.sheets.rawMycaseLeadsReport no está definido.');
  }

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clearContents();

  if (rows.length) {
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  }
}