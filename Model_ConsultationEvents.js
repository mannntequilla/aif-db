function buildConsultationEvents() {
  const bridgeLeadCaseRows = readSheetAsObjectsIfExists_(CONFIG.sheets.bridgeLeadCase);
  const rawEvents = readSheetAsObjectsIfExists_(CONFIG.sheets.rawEvents);
  const bridgeLeadCaseById = indexBridgeLeadCaseById_(bridgeLeadCaseRows);

  const rows = rawEvents
    .map(function(eventRow) {
      const caseId = String(firstNonEmpty_(extractCaseIdFromEvent_(eventRow))).trim();
      if (!caseId) return null;

      const bridgeRow = bridgeLeadCaseById[caseId];
      if (!bridgeRow) return null;

      return {
        id: firstNonEmpty_(bridgeRow.id),
        Full_name: firstNonEmpty_(bridgeRow.Full_name),
        date_added: cleanScalar_(firstNonEmpty_(bridgeRow.date_added)),
        'lead/case': firstNonEmpty_(bridgeRow['lead/case']),
        event_type: normalizeScheduledEventType_(firstNonEmpty_(eventRow.event_type, eventRow.type)),
        start: cleanScalar_(firstNonEmpty_(eventRow.start))
      };
    })
    .filter(Boolean);

  writeRowsToSheet_(CONFIG.sheets.consultationEvents, rows);
  formatConsultationEventsColumns_();
}

function indexBridgeLeadCaseById_(bridgeLeadCaseRows) {
  const out = {};

  bridgeLeadCaseRows.forEach(function(row) {
    const id = String(firstNonEmpty_(row.id)).trim();
    if (!id) return;

    out[id] = row;
  });

  return out;
}

function formatConsultationEventsColumns_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.consultationEvents);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  ['date_added'].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
  });
}
