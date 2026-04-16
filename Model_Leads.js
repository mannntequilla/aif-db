function buildFactLeads() {
  const leads = readSheetAsObjectsIfExists_(CONFIG.sheets.rawMyCaseLeadsReport);
  const events = readSheetAsObjects_(CONFIG.sheets.rawEvents);

  if (!leads || !leads.length) {
    writeRowsToSheet_(CONFIG.sheets.factLeads, []);
    formatFactLeadsColumns_();
    return;
  }

  const relevantEvents = getRelevantLeadEvents_(events);

  const rows = leads.map(function(leadRow) {
    const consultationEvent = findLatestLeadConsultationEvent_(leadRow, relevantEvents);

    return {
      lead_name: firstNonEmpty_(leadRow['Lead'], leadRow['Lead name']),
      phone_number: firstNonEmpty_(leadRow['Phone number']),
      lead_status: firstNonEmpty_(leadRow['Lead status']),
      practice_area: firstNonEmpty_(leadRow['Practice area']),
      date_added: toDateOnlyMaybe_(firstNonEmpty_(leadRow['Date added'])),
      referral_source: firstNonEmpty_(leadRow['Referral source']),
      referred_by: firstNonEmpty_(leadRow['Referred by']),
      lead_value: toNumber_(firstNonEmpty_(leadRow['Value'])),
      consultation_event_date: toDateOnlyMaybe_(firstNonEmpty_(consultationEvent && consultationEvent.event_date)),
      consultation_event_title: firstNonEmpty_(consultationEvent && consultationEvent.event_title),
      consultation_event_created_at: toDateOnlyMaybe_(firstNonEmpty_(consultationEvent && consultationEvent.event_created_at)),
      consultation_event_type: formatLeadConsultationEventType_(firstNonEmpty_(consultationEvent && consultationEvent.event_type)),
      consultation_match_score: toNumber_(consultationEvent && consultationEvent.match_score),
      has_consultation_event: consultationEvent ? 'Yes' : 'No'
    };
  });

  writeRowsToSheet_(CONFIG.sheets.factLeads, rows);
  formatFactLeadsColumns_();
}

function getRelevantLeadEvents_(events) {
  return events
    .map(function(ev) {
      const eventType = normalizeLeadEventType_(firstNonEmpty_(ev.event_type));
      if (
        eventType !== 'INITIAL CONSULTATION' &&
        eventType !== 'DETAINEE VISITATION'
      ) {
        return null;
      }

      return {
        event_type: eventType,
        event_title: firstNonEmpty_(ev.name, ev.title, ev.subject),
        event_date: firstNonEmpty_(ev.start, ev.start_at, ev.start_time, ev.date),
        event_created_at: firstNonEmpty_(ev.created_at, ev.updated_at, ev.start, ev.start_at, ev.date),
        searchable_text: normalizeText_([
          firstNonEmpty_(ev.name),
          firstNonEmpty_(ev.title),
          firstNonEmpty_(ev.subject),
          firstNonEmpty_(ev.description)
        ].join(' '))
      };
    })
    .filter(Boolean);
}

function normalizeLeadEventType_(eventType) {
  return String(firstNonEmpty_(eventType))
    .trim()
    .toUpperCase()
    .replace(/[_-]+/g, ' ');
}

function findLatestLeadConsultationEvent_(leadRow, relevantEvents) {
  const leadName = firstNonEmpty_(leadRow['Lead'], leadRow['Lead name']);
  const normalizedLeadName = normalizeText_(leadName);

  let bestMatch = null;
  let latestCreatedAt = null;

  relevantEvents.forEach(function(eventRow) {
    const matchScore = scoreLeadEventNameMatch_(eventRow.searchable_text, normalizedLeadName);
    if (matchScore < 5) return;

    const createdAt = toDateMaybe_(eventRow.event_created_at);
    const createdAtTime = createdAt && createdAt.getTime ? createdAt.getTime() : Number.NEGATIVE_INFINITY;
    const latestCreatedAtTime =
      latestCreatedAt && latestCreatedAt.getTime ? latestCreatedAt.getTime() : Number.NEGATIVE_INFINITY;

    if (!bestMatch || createdAtTime > latestCreatedAtTime) {
      bestMatch = {
        event_type: eventRow.event_type,
        event_title: eventRow.event_title,
        event_date: eventRow.event_date,
        event_created_at: eventRow.event_created_at,
        match_score: matchScore
      };
      latestCreatedAt = createdAt;
    }
  });

  return bestMatch;
}

function formatLeadConsultationEventType_(eventType) {
  const normalizedEventType = normalizeLeadEventType_(eventType);

  if (normalizedEventType === 'INITIAL CONSULTATION') {
    return 'Initial Consultation';
  }

  if (normalizedEventType === 'DETAINEE VISITATION') {
    return 'Detainee Visitation';
  }

  return '';
}

function scoreLeadEventNameMatch_(searchableText, normalizedLeadName) {
  if (!normalizedLeadName || !searchableText) return 0;
  return searchableText.indexOf(normalizedLeadName) !== -1 ? 10 : 0;
}

function formatFactLeadsColumns_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.factLeads);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  [
    'date_added',
    'consultation_event_date',
    'consultation_event_created_at'
  ].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
  });

  [
    'lead_value',
    'consultation_match_score'
  ].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('0.00');
    }
  });
}
