function buildFactCaseMaster() {
  const cases = readSheetAsObjects_(CONFIG.sheets.rawCases);
  const clients = readSheetAsObjects_(CONFIG.sheets.rawClients);
  const invoices = readSheetAsObjects_(CONFIG.sheets.rawInvoices);
  const events = readSheetAsObjects_(CONFIG.sheets.rawEvents);
  const roles = readSheetAsObjects_(CONFIG.sheets.rawRoles);
  const customFields = readSheetAsObjectsIfExists_(CONFIG.sheets.rawCustomFields);
  const mycaseLeadsReport = readSheetAsObjectsIfExists_(CONFIG.sheets.rawMyCaseLeadsReport);

  const clientsById = indexBy_(clients, 'id');
  const invoicesByCaseId = aggregateInvoicesByCaseId_(invoices);
  const firstConsultByCaseId = getFirstInitialConsultationByCaseId_(events);
  const leadMatches = buildLeadMatches_(cases, mycaseLeadsReport, clientsById);
  const retainerCustomFieldId = getCustomFieldIdByName_(customFields, 'Retainer', 'case');
  const caseTypeCustomFieldId = getCustomFieldIdByName_(customFields, 'Case Type', 'case');

  const rows = cases.map(function(caseRow) {
    const caseId = firstNonEmpty_(caseRow.id, caseRow.case_id);

    const linkedClientRef = findPreferredCaseClientRef_(caseRow);
    const linkedClient = resolveClientFromRef_(linkedClientRef, clientsById) || {};

    const financials = invoicesByCaseId[String(caseId)] || {
      total_invoice_amount: 0,
      total_paid_so_far: 0,
      total_balance: 0
    };

    const firstConsult = firstConsultByCaseId[String(caseId)] || {};
    const leadMatch = leadMatches[String(caseId)] || {};

    const caseOpenedDate = firstNonEmpty_(caseRow.opened_date, caseRow.case_opened_date);
    const firstConsultDate = firstConsult.first_initial_consultation_date || '';
    const retainerValue = getCaseCustomFieldValueById_(caseRow, retainerCustomFieldId);
    const caseTypeValue = getCaseCustomFieldValueById_(caseRow, caseTypeCustomFieldId);

    const leadType = classifyLeadType_(
      leadMatch,
      firstConsultDate,
      caseOpenedDate
    );

    const leadReferralSource = normalizeReferralSource_(
      firstNonEmpty_(leadMatch.referral_source),
      leadType
    );

    return {
      case_id: caseId,

      case_opened_date: toDateOnlyMaybe_(
        firstNonEmpty_(caseRow.opened_date, caseRow.case_opened_date)
      ),
      case_updated_at: toDateOnlyMaybe_(
        firstNonEmpty_(caseRow.updated_at, caseRow.case_updated_at)
      ),

      case_name: firstNonEmpty_(caseRow.name, caseRow.case_name),
      case_description: firstNonEmpty_(caseRow.description, caseRow.case_description),
      case_status: firstNonEmpty_(caseRow.status, caseRow.case_status),
      case_stage: firstNonEmpty_(caseRow.case_stage, caseRow.stage),
      case_type: caseTypeValue,
      practice_area: firstNonEmpty_(caseRow.practice_area, caseRow.practice_area_name),
      office_name: extractOfficeName_(caseRow),
      retainer: retainerValue,

      client_id: firstNonEmpty_(linkedClient.id, linkedClient.client_id),
      client_full_name: firstNonEmpty_(
        linkedClient.full_name,
        buildFullName_(linkedClient),
        caseRow.name
      ),
      client_email: firstNonEmpty_(linkedClient.email),
      client_phone: firstNonEmpty_(
        linkedClient.cell_phone_number,
        linkedClient.home_phone_number,
        linkedClient.work_phone_number,
        linkedClient.phone
      ),
      client_address: extractAddressLine_(linkedClient),
      client_city: extractCity_(linkedClient),
      client_state: extractState_(linkedClient),

      total_invoice_amount: financials.total_invoice_amount,
      total_paid_so_far: financials.total_paid_so_far,
      total_balance: financials.total_balance,

      first_initial_consultation_date: toDateOnlyMaybe_(
        firstConsult.first_initial_consultation_date || ''
      ),
      first_initial_consultation_title: firstConsult.first_initial_consultation_title || '',
      first_initial_consultation_event_type: firstConsult.first_initial_consultation_event_type || '',
      consultation_fee: getConsultationFeeByEventType_(
        firstConsult.first_initial_consultation_event_type || ''
      ),

      matched_lead_name: firstNonEmpty_(leadMatch.lead_name),
      matched_lead_phone_number: firstNonEmpty_(leadMatch.phone_number),
      lead_type: classifyLeadType_(
        leadMatch,
        firstConsult.first_initial_consultation_date,
        caseRow.opened_date
      ),

      lead_status: firstNonEmpty_(leadMatch.lead_status),
      lead_practice_area: firstNonEmpty_(leadMatch.practice_area),
      lead_date_added: toDateOnlyMaybe_(firstNonEmpty_(leadMatch.date_added)),
      lead_conversion_date: toDateOnlyMaybe_(firstNonEmpty_(leadMatch.conversion_date)),
      lead_referral_source: leadReferralSource,
      lead_referred_by: firstNonEmpty_(leadMatch.referred_by),
      lead_value: firstNonEmpty_(leadMatch.value),
      lead_match_method: firstNonEmpty_(leadMatch.match_method),
      lead_match_score: firstNonEmpty_(leadMatch.match_score)
    };
  });

  writeRowsToSheet_(CONFIG.sheets.factCaseMaster, rows);
  formatFactCaseMasterColumns_();
}

function buildLeadMatches_(cases, mycaseLeadsReport, clientsById) {
  const out = {};

  if (!mycaseLeadsReport || !mycaseLeadsReport.length) return out;

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

    let bestMatch = null;
    let bestScore = 0;

    mycaseLeadsReport.forEach(function(leadRow) {
      const leadName = normalizeText_(leadRow['Lead name']);
      const leadPhone = normalizePhone_(leadRow['Phone number']);
      const leadConversionDate = toDateOnlyMaybe_(leadRow['Conversion date']);

      let score = 0;

      if (caseClientName && leadName && caseClientName === leadName) score += 3;
      if (casePhone && leadPhone && casePhone === leadPhone) score += 4;
      if (caseOpenedDate && leadConversionDate && caseOpenedDate === leadConversionDate) score += 2;

      if (score > bestScore) {
        bestScore = score;
        bestMatch = {
          lead_name: leadRow['Lead name'] || '',
          lead_status: leadRow['Lead status'] || '',
          practice_area: leadRow['Practice area'] || '',
          phone_number: leadRow['Phone number'] || '',
          date_added: leadRow['Date added'] || '',
          referral_source: leadRow['Referral source'] || '',
          referred_by: leadRow['Referred by'] || '',
          value: leadRow['Value'] || '',
          conversion_date: leadRow['Conversion date'] || '',
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
      out[caseId] = bestMatch;
    }
  });

  return out;
}

function buildLeadMatchMethod_(args) {
  const parts = [];

  if (args.caseClientName && args.leadName && args.caseClientName === args.leadName) {
    parts.push('lead_name');
  }

  if (args.casePhone && args.leadPhone && args.casePhone === args.leadPhone) {
    parts.push('phone_number');
  }

  if (
    args.caseOpenedDate &&
    args.leadConversionDate &&
    args.caseOpenedDate === args.leadConversionDate
  ) {
    parts.push('conversion_date');
  }

  return parts.join('+');
}

function normalizeConsultationFeeEventType_(value) {
  return normalizeText_(String(value || '').replace(/[_-]+/g, ' '));
}

function getConsultationFeeByEventType_(eventType) {
  const normalizedEventType = normalizeConsultationFeeEventType_(eventType);

  if (normalizedEventType === normalizeConsultationFeeEventType_('Initial Consultation')) {
    return 100;
  }

  if (normalizedEventType === normalizeConsultationFeeEventType_('Detainee Visitation')) {
    return 1500;
  }

  return 0;
}

function normalizeText_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizePhone_(value) {
  const digits = String(value || '').replace(/\D+/g, '');
  if (!digits) return '';
  return digits.length > 10 ? digits.slice(-10) : digits;
}

function formatFactCaseMasterColumns_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.factCaseMaster);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  [
    'case_opened_date',
    'case_updated_at',
    'first_initial_consultation_date',
    'lead_date_added',
    'lead_conversion_date'
  ].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
  });

  [
    'total_invoice_amount',
    'total_paid_so_far',
    'total_balance',
    'consultation_fee',
    'lead_value',
    'lead_match_score'
  ].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('0.00');
    }
  });
}
