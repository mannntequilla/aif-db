function buildEventsPerCaseId() {
  const bridgeLeadCaseRows = readSheetAsObjectsIfExists_(CONFIG.sheets.bridgeLeadCase);
  const rawEvents = readSheetAsObjectsIfExists_(CONFIG.sheets.rawEvents);
  const eventsByCaseId = indexEventsByCaseId_(rawEvents);

  const rows = [];

  bridgeLeadCaseRows.forEach(function(bridgeRow) {
    const caseId = String(firstNonEmpty_(bridgeRow.id)).trim();
    const matchedEvents = caseId ? (eventsByCaseId[caseId] || []) : [];

    if (!matchedEvents.length) {
      rows.push({
        id: firstNonEmpty_(bridgeRow.id),
        Full_name: firstNonEmpty_(bridgeRow.Full_name),
        date_added: cleanScalar_(firstNonEmpty_(bridgeRow.date_added)),
        'lead/case': firstNonEmpty_(bridgeRow['lead/case']),
        event_title: 'No events were scheduled',
        event_type: 'No events were scheduled',
        start: 'No events were scheduled'
      });
      return;
    }

    matchedEvents.forEach(function(eventRow) {
      rows.push({
        id: firstNonEmpty_(bridgeRow.id),
        Full_name: firstNonEmpty_(bridgeRow.Full_name),
        date_added: cleanScalar_(firstNonEmpty_(bridgeRow.date_added)),
        'lead/case': firstNonEmpty_(bridgeRow['lead/case']),
        event_title: cleanScalar_(firstNonEmpty_(eventRow.name, eventRow.title, eventRow.subject)),
        event_type: normalizeScheduledEventType_(firstNonEmpty_(eventRow.event_type, eventRow.type)),
        start: cleanScalar_(firstNonEmpty_(eventRow.start))
      });
    });
  });

  writeRowsToSheet_(CONFIG.sheets.eventsPerCaseId, rows);
  formatEventsPerCaseIdColumns_();
}

function indexEventsByCaseId_(rawEvents) {
  const out = {};

  rawEvents.forEach(function(eventRow) {
    const caseId = String(firstNonEmpty_(extractCaseIdFromEvent_(eventRow))).trim();
    if (!caseId) return;

    if (!out[caseId]) {
      out[caseId] = [];
    }

    out[caseId].push(eventRow);
  });

  return out;
}

function formatEventsPerCaseIdColumns_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.eventsPerCaseId);
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
