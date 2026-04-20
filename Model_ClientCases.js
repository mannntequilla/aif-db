function buildBridgeClientCases() {
  const cases = readSheetAsObjectsIfExists_(CONFIG.sheets.rawCases);
  const clients = readSheetAsObjectsIfExists_(CONFIG.sheets.rawClients);
  const casesById = indexBy_(cases, 'id');
  const rows = [];

  clients.forEach(function(clientRow) {
    const clientId = String(firstNonEmpty_(clientRow.id, clientRow.client_id)).trim();
    if (!clientId) return;

    const clientCaseRefs = parseJsonMaybe_(firstNonEmpty_(clientRow.cases, '[]'));
    if (!Array.isArray(clientCaseRefs) || !clientCaseRefs.length) return;

    clientCaseRefs.forEach(function(caseRef) {
      const caseId = String(firstNonEmpty_(caseRef && caseRef.id, caseRef && caseRef.case_id)).trim();
      if (!caseId) return;

      const caseRow = casesById[caseId] || {};

      rows.push({
        client_id: clientId,
        client_full_name: firstNonEmpty_(clientRow.full_name, buildFullName_(clientRow)),
        client_created_at: toDateOnlyMaybe_(firstNonEmpty_(clientRow.created_at)),
        case_id: caseId,
        case_name: firstNonEmpty_(caseRow.name, caseRow.case_name),
        case_status: firstNonEmpty_(caseRow.status, caseRow.case_status),
        case_opened_date: toDateOnlyMaybe_(firstNonEmpty_(caseRow.opened_date, caseRow.case_opened_date))
      });
    });
  });

  writeRowsToSheet_(CONFIG.sheets.bridgeClientCases, rows);
  formatBridgeClientCasesColumns_();
}

function formatBridgeClientCasesColumns_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.bridgeClientCases);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  ['client_created_at', 'case_opened_date'].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
  });
}
