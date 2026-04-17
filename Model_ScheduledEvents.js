function buildFactScheduledEvents() {
  const events = readSheetAsObjects_(CONFIG.sheets.rawEvents);
  const leads = readSheetAsObjectsIfExists_(CONFIG.sheets.rawLeads);
  const cases = readSheetAsObjectsIfExists_(CONFIG.sheets.rawCases);
  const clients = readSheetAsObjectsIfExists_(CONFIG.sheets.rawClients);

  if (!events || !events.length) {
    writeRowsToSheet_(CONFIG.sheets.factScheduledEvents, []);
    formatFactScheduledEventsColumns_();
    return;
  }

  const leadsByCaseId = indexLeadsByCaseId_(leads);
  const casesById = indexBy_(cases, 'id');
  const clientsById = indexBy_(clients, 'id');

  const rows = events.map(function(eventRow) {
    const eventCaseId = String(firstNonEmpty_(eventRow.case_id, safeGet_(parseJsonMaybe_(eventRow.case), 'id', '')) || '').trim();
    const matchedLead = eventCaseId ? leadsByCaseId[eventCaseId] : null;
    const matchedCase = !matchedLead && eventCaseId ? casesById[eventCaseId] : null;
    const recordStage = matchedLead ? 'Lead' : (matchedCase ? 'Case' : 'Unknown');
    const associatedClient = matchedCase
      ? resolveClientFromRef_(findPreferredCaseClientRef_(matchedCase), clientsById) || {}
      : {};

    return {
      event_id: firstNonEmpty_(eventRow.id, eventRow.event_id),
      case_id: eventCaseId,
      event_title: firstNonEmpty_(eventRow.name, eventRow.title, eventRow.subject),
      event_type: normalizeScheduledEventType_(firstNonEmpty_(eventRow.event_type, eventRow.type)),
      event_start: toDateMaybe_(firstNonEmpty_(eventRow.start)),
      event_end: toDateMaybe_(firstNonEmpty_(eventRow.end)),
      dim_date: toDateOnlyMaybe_(firstNonEmpty_(eventRow.start)),
      record_stage: recordStage,
      record_name: matchedLead
        ? extractLeadDisplayName_(matchedLead)
        : matchedCase
          ? firstNonEmpty_(matchedCase.name, matchedCase.case_name)
          : '',
      associated_contact_name: matchedCase
        ? firstNonEmpty_(associatedClient.full_name, buildFullName_(associatedClient))
        : '',
      event_created_at: toDateMaybe_(firstNonEmpty_(eventRow.created_at)),
      event_updated_at: toDateMaybe_(firstNonEmpty_(eventRow.updated_at))
    };
  });

  writeRowsToSheet_(CONFIG.sheets.factScheduledEvents, rows);
  formatFactScheduledEventsColumns_();
}

function indexLeadsByCaseId_(leads) {
  const out = {};

  leads.forEach(function(leadRow) {
    const leadCaseId = String(
      firstNonEmpty_(
        leadRow.case_id,
        safeGet_(parseJsonMaybe_(leadRow.case), 'id', ''),
        leadRow.case
      ) || ''
    ).trim();

    if (!leadCaseId) return;
    out[leadCaseId] = leadRow;
  });

  return out;
}

function extractLeadDisplayName_(leadRow) {
  return firstNonEmpty_(
    leadRow.name,
    leadRow.lead_name,
    leadRow.full_name,
    buildFullName_(leadRow),
    [firstNonEmpty_(leadRow.first_name), firstNonEmpty_(leadRow.last_name)].filter(Boolean).join(' ').trim()
  );
}

function normalizeScheduledEventType_(eventType) {
  return String(firstNonEmpty_(eventType))
    .trim()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ');
}

function formatFactScheduledEventsColumns_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.factScheduledEvents);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  [
    'dim_date',
    'event_start',
    'event_end',
    'event_created_at',
    'event_updated_at'
  ].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      const format = name === 'dim_date' ? 'yyyy-mm-dd' : 'yyyy-mm-dd hh:mm:ss';
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat(format);
    }
  });
}
