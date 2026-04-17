function buildFactContactSchedules() {
  const scheduledEvents = readSheetAsObjectsIfExists_(CONFIG.sheets.factScheduledEvents);
  const leads = readSheetAsObjectsIfExists_(CONFIG.sheets.rawLeads);

  const rowsByNameKey = {};

  leads.forEach(function(leadRow) {
    const fullName = extractRawLeadFullName_(leadRow);
    const nameKey = normalizeText_(fullName);
    if (!nameKey) return;

    if (!rowsByNameKey[nameKey]) {
      rowsByNameKey[nameKey] = {
        contact_name: fullName,
        scheduled_event_title: '',
        scheduled_event_type: '',
        scheduled_event_start: '',
        scheduled_event_record_stage: '',
        has_scheduled_event: 'No'
      };
    }
  });

  scheduledEvents.forEach(function(eventRow) {
    const contactName = String(firstNonEmpty_(eventRow.associated_contact_name)).trim();
    const nameKey = normalizeText_(contactName);
    if (!nameKey) return;

    if (!rowsByNameKey[nameKey]) {
      rowsByNameKey[nameKey] = {
        contact_name: contactName,
        scheduled_event_title: '',
        scheduled_event_type: '',
        scheduled_event_start: '',
        scheduled_event_record_stage: '',
        has_scheduled_event: 'No'
      };
    }

    const currentRow = rowsByNameKey[nameKey];
    const candidateStart = toDateMaybe_(firstNonEmpty_(eventRow.event_start, eventRow.event_created_at));
    const currentStart = toDateMaybe_(firstNonEmpty_(currentRow.scheduled_event_start));

    const candidateTime = candidateStart && candidateStart.getTime ? candidateStart.getTime() : Number.NEGATIVE_INFINITY;
    const currentTime = currentStart && currentStart.getTime ? currentStart.getTime() : Number.NEGATIVE_INFINITY;

    if (candidateTime >= currentTime) {
      currentRow.contact_name = contactName;
      currentRow.scheduled_event_title = firstNonEmpty_(eventRow.event_title);
      currentRow.scheduled_event_type = firstNonEmpty_(eventRow.event_type);
      currentRow.scheduled_event_start = toDateMaybe_(firstNonEmpty_(eventRow.event_start));
      currentRow.scheduled_event_record_stage = firstNonEmpty_(eventRow.record_stage);
      currentRow.has_scheduled_event = firstNonEmpty_(eventRow.event_title, eventRow.event_type) ? 'Yes' : 'No';
    }
  });

  const rows = Object.keys(rowsByNameKey)
    .sort()
    .map(function(nameKey) {
      return rowsByNameKey[nameKey];
    });

  writeRowsToSheet_(CONFIG.sheets.factLeads, rows);
  formatFactContactSchedulesColumns_();
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

function formatFactContactSchedulesColumns_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.factLeads);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  ['scheduled_event_start'].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    }
  });
}
