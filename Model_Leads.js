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
    const initialConsultationEvent = findLatestLeadEventByType_(
      leadRow,
      relevantEvents,
      'INITIAL CONSULTATION'
    );
    const detaineeVisitationEvent = findLatestLeadEventByType_(
      leadRow,
      relevantEvents,
      'DETAINEE VISITATION'
    );

    return {
      lead_name: firstNonEmpty_(leadRow['Lead'], leadRow['Lead name']),
      phone_number: firstNonEmpty_(leadRow['Phone number']),
      lead_status: firstNonEmpty_(leadRow['Lead status']),
      practice_area: firstNonEmpty_(leadRow['Practice area']),
      date_added: toDateOnlyMaybe_(firstNonEmpty_(leadRow['Date added'])),
      lead_month: formatCaseProfitabilityEntryMonth_(firstNonEmpty_(leadRow['Date added'])),
      conversion_date: toDateOnlyMaybe_(firstNonEmpty_(leadRow['Conversion date'])),
      referral_source: firstNonEmpty_(leadRow['Referral source']),
      referred_by: firstNonEmpty_(leadRow['Referred by']),
      lead_value: toNumber_(firstNonEmpty_(leadRow['Value'])),

      initial_consultation_event_date: toDateOnlyMaybe_(firstNonEmpty_(initialConsultationEvent && initialConsultationEvent.event_date)),
      initial_consultation_event_title: firstNonEmpty_(initialConsultationEvent && initialConsultationEvent.event_title),
      initial_consultation_event_created_at: toDateOnlyMaybe_(firstNonEmpty_(initialConsultationEvent && initialConsultationEvent.event_created_at)),
      initial_consultation_match_score: toNumber_(initialConsultationEvent && initialConsultationEvent.match_score),

      detainee_visitation_event_date: toDateOnlyMaybe_(firstNonEmpty_(detaineeVisitationEvent && detaineeVisitationEvent.event_date)),
      detainee_visitation_event_title: firstNonEmpty_(detaineeVisitationEvent && detaineeVisitationEvent.event_title),
      detainee_visitation_event_created_at: toDateOnlyMaybe_(firstNonEmpty_(detaineeVisitationEvent && detaineeVisitationEvent.event_created_at)),
      detainee_visitation_match_score: toNumber_(detaineeVisitationEvent && detaineeVisitationEvent.match_score),

      has_initial_consultation_event: initialConsultationEvent ? 'Yes' : 'No',
      has_detainee_visitation_event: detaineeVisitationEvent ? 'Yes' : 'No'
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

function findLatestLeadEventByType_(leadRow, relevantEvents, eventType) {
  const leadName = firstNonEmpty_(leadRow['Lead'], leadRow['Lead name']);
  const normalizedLeadName = normalizeText_(leadName);
  const leadNameParts = splitLeadNameParts_(leadName);

  let bestMatch = null;
  let latestCreatedAt = null;

  relevantEvents.forEach(function(eventRow) {
    if (eventRow.event_type !== eventType) return;

    const matchScore = scoreLeadEventNameMatch_(eventRow.searchable_text, normalizedLeadName, leadNameParts);
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

function splitLeadNameParts_(leadName) {
  const cleanName = normalizeText_(leadName);
  const parts = cleanName.split(' ').filter(Boolean);

  return {
    full_name: cleanName,
    first_name: parts[0] || '',
    last_name: parts.length ? parts[parts.length - 1] : ''
  };
}

function scoreLeadEventNameMatch_(searchableText, normalizedLeadName, leadNameParts) {
  let score = 0;

  if (normalizedLeadName && searchableText.indexOf(normalizedLeadName) !== -1) {
    score += 6;
  }

  if (leadNameParts.first_name && searchableText.indexOf(leadNameParts.first_name) !== -1) {
    score += 2;
  }

  if (leadNameParts.last_name && searchableText.indexOf(leadNameParts.last_name) !== -1) {
    score += 3;
  }

  return score;
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
    'conversion_date',
    'initial_consultation_event_date',
    'initial_consultation_event_created_at',
    'detainee_visitation_event_date',
    'detainee_visitation_event_created_at'
  ].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
  });

  [
    'lead_value',
    'initial_consultation_match_score',
    'detainee_visitation_match_score'
  ].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('0.00');
    }
  });
}
