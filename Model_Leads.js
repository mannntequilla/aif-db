function buildFactLeads() {
  const leads = readSheetAsObjectsIfExists_(CONFIG.sheets.rawMyCaseLeadsReport);
  const events = readSheetAsObjects_(CONFIG.sheets.rawEvents);

  if (!leads || !leads.length) {
    writeRowsToSheet_(CONFIG.sheets.factLeads, []);
    formatFactLeadsColumns_();
    return;
  }

  const relevantEvents = getRelevantConsultationEvents_(events);

  const rows = leads.map(function(leadRow) {
    const leadStatus = firstNonEmpty_(leadRow['Lead status']);
    const inferredCategory = getConsultationCategoryByLeadStatus_(leadStatus);
    const shouldLookupConsultation = !!inferredCategory;
    const matchedEvent = shouldLookupConsultation
      ? findBestConsultationEventForLead_(leadRow, relevantEvents, inferredCategory)
      : null;

    const eventType = firstNonEmpty_(matchedEvent && matchedEvent.event_type);

    return {
      lead_name: firstNonEmpty_(leadRow['Lead'], leadRow['Lead name']),
      phone_number: firstNonEmpty_(leadRow['Phone number']),
      lead_status: leadStatus,
      practice_area: firstNonEmpty_(leadRow['Practice area']),
      date_added: toDateOnlyMaybe_(firstNonEmpty_(leadRow['Date added'])),
      lead_month: formatCaseProfitabilityEntryMonth_(firstNonEmpty_(leadRow['Date added'])),
      conversion_date: toDateOnlyMaybe_(firstNonEmpty_(leadRow['Conversion date'])),
      referral_source: firstNonEmpty_(leadRow['Referral source']),
      referred_by: firstNonEmpty_(leadRow['Referred by']),
      lead_value: toNumber_(firstNonEmpty_(leadRow['Value'])),
      consultation_lookup_required: shouldLookupConsultation ? 'Yes' : 'No',
      consultation_event_date: toDateOnlyMaybe_(firstNonEmpty_(matchedEvent && matchedEvent.event_date)),
      consultation_event_type: eventType,
      consultation_event_title: firstNonEmpty_(matchedEvent && matchedEvent.event_title),
      consultation_event_match_score: toNumber_(matchedEvent && matchedEvent.match_score),
      consultation_category: firstNonEmpty_(
        getConsultationCategoryByEventType_(eventType),
        inferredCategory
      ),
      has_consultation_event: firstNonEmpty_(
        eventType ? 'Yes' : '',
        inferredCategory ? 'Yes' : '',
        'No'
      )
    };
  });

  writeRowsToSheet_(CONFIG.sheets.factLeads, rows);
  formatFactLeadsColumns_();
}

function getLeadStatusesThatImplyInitialConsultation_() {
  return [
    'consult scheduled',
    'no show',
    'first follow-up',
    'second follow-up',
    'hot deal',
    'contract',
    'no case'
  ];
}

function getLeadStatusesThatImplyDetaineeVisitation_() {
  return [
    'direct visitation',
    'detainee visitation'
  ];
}

function getConsultationCategoryByLeadStatus_(status) {
  const normalizedStatus = normalizeLeadStatus_(status);

  if (getLeadStatusesThatImplyDetaineeVisitation_().indexOf(normalizedStatus) !== -1) {
    return 'Detainee Visitation';
  }

  if (getLeadStatusesThatImplyInitialConsultation_().indexOf(normalizedStatus) !== -1) {
    return 'Initial Consultation';
  }

  return '';
}

function getRelevantConsultationEvents_(events) {
  return events
    .map(function(ev) {
      const eventType = String(firstNonEmpty_(ev.event_type)).trim().toUpperCase();
      if (
        eventType !== 'INITIAL CONSULTATION' &&
        eventType !== 'DETAINEE VISITATION'
      ) {
        return null;
      }

      const eventTitle = firstNonEmpty_(ev.name, ev.title, ev.subject);
      const eventDate = firstNonEmpty_(ev.start, ev.start_at, ev.start_time, ev.date);
      const searchableText = normalizeText_([
        firstNonEmpty_(ev.name),
        firstNonEmpty_(ev.title),
        firstNonEmpty_(ev.subject),
        firstNonEmpty_(ev.description)
      ].join(' '));

      return {
        event_type: eventType,
        event_title: eventTitle,
        event_date: eventDate,
        searchable_text: searchableText
      };
    })
    .filter(Boolean);
}

function findBestConsultationEventForLead_(leadRow, relevantEvents, expectedCategory) {
  const leadName = firstNonEmpty_(leadRow['Lead'], leadRow['Lead name']);
  const leadDateAdded = toDateOnlyMaybe_(firstNonEmpty_(leadRow['Date added']));
  const nameParts = splitLeadNameParts_(leadName);
  const normalizedFullName = normalizeText_(leadName);

  let bestMatch = null;
  let bestScore = 0;

  relevantEvents.forEach(function(eventRow) {
    if (
      expectedCategory &&
      getConsultationCategoryByEventType_(eventRow.event_type) !== expectedCategory
    ) {
      return;
    }

    const score = scoreConsultationEventMatch_(eventRow, normalizedFullName, nameParts, leadDateAdded);
    if (score > bestScore) {
      bestScore = score;
      bestMatch = {
        event_type: eventRow.event_type,
        event_title: eventRow.event_title,
        event_date: eventRow.event_date,
        match_score: score
      };
    }
  });

  return bestScore >= 5 ? bestMatch : null;
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

function scoreConsultationEventMatch_(eventRow, normalizedFullName, nameParts, leadDateAdded) {
  const searchableText = eventRow.searchable_text || '';
  let score = 0;

  if (normalizedFullName && searchableText.indexOf(normalizedFullName) !== -1) {
    score += 6;
  }

  if (nameParts.first_name && searchableText.indexOf(nameParts.first_name) !== -1) {
    score += 2;
  }

  if (nameParts.last_name && searchableText.indexOf(nameParts.last_name) !== -1) {
    score += 3;
  }

  const eventDate = toDateOnlyMaybe_(eventRow.event_date);
  if (leadDateAdded && eventDate) {
    const dayDiff = Math.abs((eventDate.getTime() - leadDateAdded.getTime()) / 86400000);

    if (dayDiff <= 7) score += 3;
    else if (dayDiff <= 30) score += 1;
  }

  return score;
}

function getConsultationCategoryByEventType_(eventType) {
  const normalizedEventType = normalizeConsultationFeeEventType_(eventType);

  if (normalizedEventType === normalizeConsultationFeeEventType_('Initial Consultation')) {
    return 'Initial Consultation';
  }

  if (normalizedEventType === normalizeConsultationFeeEventType_('Detainee Visitation')) {
    return 'Detainee Visitation';
  }

  return '';
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
    'consultation_event_date'
  ].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
  });

  ['lead_value', 'consultation_event_match_score'].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('0.00');
    }
  });
}
