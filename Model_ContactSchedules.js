function buildFactContactSchedules() {
  const scheduledEvents = readSheetAsObjectsIfExists_(CONFIG.sheets.factScheduledEvents);
  const leads = readSheetAsObjectsIfExists_(CONFIG.sheets.rawLeads);
  const cases = readSheetAsObjectsIfExists_(CONFIG.sheets.rawCases);
  const clients = readSheetAsObjectsIfExists_(CONFIG.sheets.rawClients);

  const casesById = indexBy_(cases, 'id');
  const clientsById = indexBy_(clients, 'id');
  const rowsByCaseId = {};

  leads.forEach(function(leadRow) {
    const caseId = extractLeadCaseId_(leadRow);
    const contactName = extractRawLeadFullName_(leadRow);

    if (!caseId || !contactName) return;

    rowsByCaseId[caseId] = {
      date_added: toDateOnlyMaybe_(firstNonEmpty_(leadRow.created_at)),
      case_id: caseId,
      contact_name: contactName,
      contact_stage: 'Lead',
      event_title: 'No event found',
      event_start: '',
      event_type: ''
    };
  });

  scheduledEvents.forEach(function(eventRow) {
    const caseId = String(firstNonEmpty_(eventRow.case_id)).trim();
    if (!caseId) return;

    const existingRow = rowsByCaseId[caseId];
    const contactName = String(firstNonEmpty_(eventRow.associated_contact_name, eventRow.record_name)).trim();
    const candidateEventStart = toDateMaybe_(firstNonEmpty_(eventRow.event_start, eventRow.event_created_at));
    const currentEventStart = existingRow
      ? toDateMaybe_(firstNonEmpty_(existingRow.event_start))
      : null;

    if (!rowsByCaseId[caseId]) {
      rowsByCaseId[caseId] = {
        date_added: resolveCaseContactCreatedAt_(caseId, casesById, clientsById),
        case_id: caseId,
        contact_name: contactName,
        contact_stage: firstNonEmpty_(eventRow.record_stage),
        event_title: firstNonEmpty_(eventRow.event_title, 'No event found'),
        event_start: toDateMaybe_(firstNonEmpty_(eventRow.event_start)),
        event_type: firstNonEmpty_(eventRow.event_type)
      };
      return;
    }

    if (!rowsByCaseId[caseId].contact_name && contactName) {
      rowsByCaseId[caseId].contact_name = contactName;
    }

    if (!rowsByCaseId[caseId].contact_stage) {
      rowsByCaseId[caseId].contact_stage = firstNonEmpty_(eventRow.record_stage);
    }

    const candidateTime = candidateEventStart && candidateEventStart.getTime
      ? candidateEventStart.getTime()
      : Number.NEGATIVE_INFINITY;
    const currentTime = currentEventStart && currentEventStart.getTime
      ? currentEventStart.getTime()
      : Number.NEGATIVE_INFINITY;

    if (candidateTime >= currentTime && firstNonEmpty_(eventRow.event_title, eventRow.event_type)) {
      rowsByCaseId[caseId].event_title = firstNonEmpty_(eventRow.event_title, 'No event found');
      rowsByCaseId[caseId].event_start = toDateMaybe_(firstNonEmpty_(eventRow.event_start));
      rowsByCaseId[caseId].event_type = firstNonEmpty_(eventRow.event_type);
    }
  });

  const rows = Object.keys(rowsByCaseId)
    .sort(function(a, b) {
      return String(a).localeCompare(String(b));
    })
    .map(function(caseId) {
      return rowsByCaseId[caseId];
    });

  writeRowsToSheet_(CONFIG.sheets.factLeads, rows);
  formatFactContactSchedulesColumns_();
}

function extractLeadCaseId_(leadRow) {
  return String(
    firstNonEmpty_(
      leadRow.case_id,
      safeGet_(parseJsonMaybe_(leadRow.case), 'id', ''),
      leadRow.case
    ) || ''
  ).trim();
}

function extractRawLeadFullName_(leadRow) {
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

function resolveCaseContactCreatedAt_(caseId, casesById, clientsById) {
  const caseRow = casesById[String(caseId)] || null;
  if (!caseRow) return '';

  const clientRef = findPreferredCaseClientRef_(caseRow);
  const client = resolveClientFromRef_(clientRef, clientsById);

  return client ? toDateOnlyMaybe_(firstNonEmpty_(client.created_at)) : '';
}

function formatFactContactSchedulesColumns_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.factLeads);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  ['date_added', 'event_start'].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      const format = name === 'date_added' ? 'yyyy-mm-dd' : 'yyyy-mm-dd hh:mm:ss';
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat(format);
    }
  });
}
