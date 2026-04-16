function buildFactLeads() {
  const leads = readSheetAsObjectsIfExists_(CONFIG.sheets.rawMyCaseLeadsReport);

  if (!leads || !leads.length) {
    writeRowsToSheet_(CONFIG.sheets.factLeads, []);
    formatFactLeadsColumns_();
    return;
  }

  const rows = leads.map(function(leadRow) {
    const leadStatus = firstNonEmpty_(leadRow['Lead status']);
    const consultationEventType = classifyLeadConsultationEventTypeByStatus_(leadStatus);

    return {
      lead_name: firstNonEmpty_(leadRow['Lead'], leadRow['Lead name']),
      phone_number: firstNonEmpty_(leadRow['Phone number']),
      lead_status: leadStatus,
      practice_area: firstNonEmpty_(leadRow['Practice area']),
      date_added: toDateOnlyMaybe_(firstNonEmpty_(leadRow['Date added'])),
      referral_source: firstNonEmpty_(leadRow['Referral source']),
      referred_by: firstNonEmpty_(leadRow['Referred by']),
      lead_value: toNumber_(firstNonEmpty_(leadRow['Value'])),
      consultation_event_date: '',
      consultation_event_title: '',
      consultation_event_created_at: '',
      consultation_event_type: consultationEventType,
      consultation_match_score: '',
      has_consultation_event: consultationEventType ? 'Yes' : 'No'
    };
  });

  writeRowsToSheet_(CONFIG.sheets.factLeads, rows);
  formatFactLeadsColumns_();
}

function getLeadStatusesForInitialConsultation_() {
  return [
    'consult scheduled',
    'first follow-up',
    'second follow-up',
    'hot deal',
    'no show',
    'no case'
  ];
}

function getLeadStatusesForDetaineeVisitation_() {
  return [
    'detainee visitation'
  ];
}

function classifyLeadConsultationEventTypeByStatus_(status) {
  const normalizedStatus = normalizeLeadStatus_(status);

  if (getLeadStatusesForDetaineeVisitation_().indexOf(normalizedStatus) !== -1) {
    return 'Detainee Visitation';
  }

  if (getLeadStatusesForInitialConsultation_().indexOf(normalizedStatus) !== -1) {
    return 'Initial Consultation';
  }

  return '';
}

function normalizeLeadStatus_(status) {
  return String(status || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
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
