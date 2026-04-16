function buildFactLeads() {
  const leads = readSheetAsObjectsIfExists_(CONFIG.sheets.rawMyCaseLeadsReport);
  const cases = readSheetAsObjects_(CONFIG.sheets.rawCases);
  const clients = readSheetAsObjects_(CONFIG.sheets.rawClients);
  const events = readSheetAsObjects_(CONFIG.sheets.rawEvents);

  if (!leads || !leads.length) {
    writeRowsToSheet_(CONFIG.sheets.factLeads, []);
    formatFactLeadsColumns_();
    return;
  }

  const clientsById = indexBy_(clients, 'id');
  const firstConsultByCaseId = getFirstInitialConsultationByCaseId_(events);
  const caseMatchesByLeadIndex = buildCaseMatchesByLeadIndex_(leads, cases, clientsById);

  const rows = leads.map(function(leadRow, leadIndex) {
    const matchedCase = caseMatchesByLeadIndex[leadIndex] || {};
    const caseId = firstNonEmpty_(matchedCase.case_id);
    const consultationEvent = caseId ? (firstConsultByCaseId[String(caseId)] || {}) : {};
    const eventType = firstNonEmpty_(consultationEvent.first_initial_consultation_event_type);

    return {
      lead_name: firstNonEmpty_(leadRow['Lead name']),
      phone_number: firstNonEmpty_(leadRow['Phone number']),
      lead_status: firstNonEmpty_(leadRow['Lead status']),
      practice_area: firstNonEmpty_(leadRow['Practice area']),
      date_added: toDateOnlyMaybe_(firstNonEmpty_(leadRow['Date added'])),
      lead_month: formatCaseProfitabilityEntryMonth_(firstNonEmpty_(leadRow['Date added'])),
      conversion_date: toDateOnlyMaybe_(firstNonEmpty_(leadRow['Conversion date'])),
      referral_source: firstNonEmpty_(leadRow['Referral source']),
      referred_by: firstNonEmpty_(leadRow['Referred by']),
      lead_value: toNumber_(firstNonEmpty_(leadRow['Value'])),
      matched_case_id: caseId,
      matched_case_name: firstNonEmpty_(matchedCase.case_name),
      matched_case_opened_date: toDateOnlyMaybe_(firstNonEmpty_(matchedCase.case_opened_date)),
      lead_match_score: firstNonEmpty_(matchedCase.match_score),
      lead_match_method: firstNonEmpty_(matchedCase.match_method),
      consultation_event_date: toDateOnlyMaybe_(firstNonEmpty_(consultationEvent.first_initial_consultation_date)),
      consultation_event_type: eventType,
      consultation_category: getConsultationCategoryByEventType_(eventType),
      has_consultation_event: eventType ? 'Yes' : 'No'
    };
  });

  writeRowsToSheet_(CONFIG.sheets.factLeads, rows);
  formatFactLeadsColumns_();
}

function buildCaseMatchesByLeadIndex_(leads, cases, clientsById) {
  const out = {};

  leads.forEach(function(leadRow, leadIndex) {
    const leadName = normalizeText_(leadRow['Lead name']);
    const leadPhone = normalizePhone_(leadRow['Phone number']);
    const leadConversionDate = toDateOnlyMaybe_(leadRow['Conversion date']);

    let bestMatch = null;
    let bestScore = 0;

    cases.forEach(function(caseRow) {
      const caseId = String(firstNonEmpty_(caseRow.id, caseRow.case_id) || '');
      if (!caseId) return;

      const linkedClientRef = findPreferredCaseClientRef_(caseRow);
      const linkedClient = resolveClientFromRef_(linkedClientRef, clientsById) || {};

      const caseOpenedDate = toDateOnlyMaybe_(
        firstNonEmpty_(caseRow.opened_date, caseRow.case_opened_date)
      );

      const caseClientName = normalizeText_(
        firstNonEmpty_(
          linkedClient.full_name,
          buildFullName_(linkedClient),
          caseRow.name
        )
      );

      const casePhone = normalizePhone_(
        firstNonEmpty_(
          linkedClient.cell_phone_number,
          linkedClient.home_phone_number,
          linkedClient.work_phone_number,
          linkedClient.phone
        )
      );

      let score = 0;

      if (caseClientName && leadName && caseClientName === leadName) score += 3;
      if (casePhone && leadPhone && casePhone === leadPhone) score += 4;
      if (caseOpenedDate && leadConversionDate && caseOpenedDate === leadConversionDate) score += 2;

      if (score > bestScore) {
        bestScore = score;
        bestMatch = {
          case_id: caseId,
          case_name: firstNonEmpty_(caseRow.name, caseRow.case_name),
          case_opened_date: firstNonEmpty_(caseRow.opened_date, caseRow.case_opened_date),
          match_score: score,
          match_method: buildLeadMatchMethod_({
            caseClientName: caseClientName,
            leadName: leadName,
            casePhone: casePhone,
            leadPhone: leadPhone,
            caseOpenedDate: caseOpenedDate,
            leadConversionDate: leadConversionDate
          })
        };
      }
    });

    if (bestMatch && bestScore >= 4) {
      out[leadIndex] = bestMatch;
    }
  });

  return out;
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
    'matched_case_opened_date',
    'consultation_event_date'
  ].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
  });

  ['lead_match_score', 'lead_value'].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('0.00');
    }
  });
}
