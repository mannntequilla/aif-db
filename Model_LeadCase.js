function buildBridgeLeadCase() {
  const rawLeads = readSheetAsObjectsIfExists_(CONFIG.sheets.rawLeads);
  const bridgeClientCases = readSheetAsObjectsIfExists_(CONFIG.sheets.bridgeClientCases);
  const rows = []
    .concat(buildBridgeLeadCaseRowsFromLeads_(rawLeads))
    .concat(buildBridgeLeadCaseRowsFromClientCases_(bridgeClientCases));

  writeRowsToSheet_(CONFIG.sheets.bridgeLeadCase, rows);
  formatBridgeLeadCaseColumns_();
}

function buildBridgeLeadCaseRowsFromLeads_(rawLeads) {
  return rawLeads
    .map(function(leadRow) {
      const leadCaseId = extractLeadCaseIdForBridge_(leadRow);
      const fullName = extractLeadFullNameForBridge_(leadRow);

      if (!leadCaseId || !fullName) return null;

      return {
        id: leadCaseId,
        Full_name: fullName,
        date_added: toDateOnlyMaybe_(firstNonEmpty_(leadRow.created_at)),
        Status: firstNonEmpty_(leadRow.status, leadRow.Status),
        'lead/case': 'Lead',
        referral_source: firstNonEmpty_(leadRow.referral_source, leadRow['Referral source'])
      };
    })
    .filter(Boolean);
}

function buildBridgeLeadCaseRowsFromClientCases_(bridgeClientCases) {
  return bridgeClientCases
    .map(function(clientCaseRow) {
      const caseId = String(firstNonEmpty_(clientCaseRow.case_id)).trim();
      const fullName = String(firstNonEmpty_(clientCaseRow.client_full_name)).trim();

      if (!caseId || !fullName) return null;

      return {
        id: caseId,
        Full_name: fullName,
        date_added: toDateOnlyMaybe_(firstNonEmpty_(clientCaseRow.client_created_at)),
        Status: firstNonEmpty_(clientCaseRow.case_stage),
        'lead/case': 'Case',
        referral_source: ''
      };
    })
    .filter(Boolean);
}

function extractLeadCaseIdForBridge_(leadRow) {
  return String(
    firstNonEmpty_(
      leadRow.case_id,
      safeGet_(parseJsonMaybe_(leadRow.case), 'id', ''),
      leadRow.case
    ) || ''
  ).trim();
}

function extractLeadFullNameForBridge_(leadRow) {
  return String(
    firstNonEmpty_(
      leadRow.full_name,
      leadRow.name,
      leadRow.lead_name,
      [
        firstNonEmpty_(leadRow['First Name'], leadRow.first_name),
        firstNonEmpty_(leadRow['Middle Name'], leadRow.middle_name),
        firstNonEmpty_(leadRow['Last Name'], leadRow.last_name)
      ].filter(Boolean).join(' ').trim()
    )
  ).trim();
}

function formatBridgeLeadCaseColumns_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.bridgeLeadCase);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const col = headers.indexOf('date_added') + 1;

  if (col > 0) {
    sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
  }
}
