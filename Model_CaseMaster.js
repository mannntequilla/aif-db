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
      case_created_at: toDateOnlyMaybe_(
        firstNonEmpty_(caseRow.created_at, caseRow.case_created_at)
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
      lead_match_score: firstNonEmpty_(leadMatch.match_score),
      lead_match_confidence: firstNonEmpty_(leadMatch.match_confidence),
      lead_match_status: firstNonEmpty_(leadMatch.match_status, 'unmatched'),
      lead_match_candidate_count: firstNonEmpty_(leadMatch.match_candidate_count, 0)
    };
  });

  writeRowsToSheet_(CONFIG.sheets.factCaseMaster, rows);
  formatFactCaseMasterColumns_();
}

function buildLeadMatches_(cases, mycaseLeadsReport, clientsById) {
  const out = {};

  if (!mycaseLeadsReport || !mycaseLeadsReport.length) return out;

  const candidates = [];

  cases.forEach(function(caseRow) {
    const caseId = String(firstNonEmpty_(caseRow.id, caseRow.case_id) || '');
    if (!caseId) return;

    const linkedClientRef = findPreferredCaseClientRef_(caseRow);
    const linkedClient = resolveClientFromRef_(linkedClientRef, clientsById) || {};

    const caseCandidateDates = getCaseLeadMatchDates_(caseRow);

    const caseClientName = normalizeText_(
      firstNonEmpty_(
        linkedClient.full_name,
        buildFullName_(linkedClient),
        caseRow.name
      )
    );

    const casePracticeArea = normalizeText_(
      firstNonEmpty_(caseRow.practice_area, caseRow.practice_area_name)
    );

    mycaseLeadsReport.forEach(function(leadRow, reportIndex) {
      // Current MyCase Leads Referral Source export header is "Lead". The
      // fallback preserves compatibility with older exports that used "Lead name".
      const leadName = normalizeText_(
        firstNonEmpty_(leadRow['Lead'], leadRow['Lead name'])
      );
      const leadConversionDate = toDateOnlyKey_(leadRow['Conversion date']);
      const leadPracticeArea = normalizeText_(leadRow['Practice area']);

      const isNameMatch = Boolean(
        caseClientName && leadName && caseClientName === leadName
      );
      const isConversionDateMatch = Boolean(
        leadConversionDate && caseCandidateDates.indexOf(leadConversionDate) !== -1
      );
      const hasComparablePracticeArea = Boolean(casePracticeArea && leadPracticeArea);
      const isPracticeAreaMatch = !hasComparablePracticeArea ||
        casePracticeArea === leadPracticeArea;

      // The report has no MyCase IDs, email, or phone number. Do not generate
      // an attribution from a name-only match; it must also agree on conversion
      // date. A populated but different practice area is a hard rejection.
      if (!isNameMatch || !isConversionDateMatch || !isPracticeAreaMatch) return;

      candidates.push({
        case_id: caseId,
        report_key: String(reportIndex),
        lead_name: firstNonEmpty_(leadRow['Lead'], leadRow['Lead name']),
          lead_status: leadRow['Lead status'] || '',
          practice_area: leadRow['Practice area'] || '',
          date_added: leadRow['Date added'] || '',
          referral_source: leadRow['Referral source'] || '',
          referred_by: leadRow['Referred by'] || '',
          value: leadRow['Value'] || '',
          conversion_date: leadRow['Conversion date'] || '',
        match_score: hasComparablePracticeArea ? 10 : 8,
        match_method: hasComparablePracticeArea
          ? 'lead_name+conversion_date+practice_area'
          : 'lead_name+conversion_date',
        match_confidence: hasComparablePracticeArea ? 'high' : 'medium',
        match_status: 'candidate'
      });
    });
  });

  const candidateCountByCaseId = {};
  const candidateCountByReportKey = {};

  candidates.forEach(function(candidate) {
    candidateCountByCaseId[candidate.case_id] =
      (candidateCountByCaseId[candidate.case_id] || 0) + 1;
    candidateCountByReportKey[candidate.report_key] =
      (candidateCountByReportKey[candidate.report_key] || 0) + 1;
  });

  Object.keys(candidateCountByCaseId).forEach(function(caseId) {
    const caseCandidates = candidates.filter(function(candidate) {
      return candidate.case_id === caseId;
    });
    const hasAmbiguousReportRow = caseCandidates.some(function(candidate) {
      return candidateCountByReportKey[candidate.report_key] !== 1;
    });

    if (candidateCountByCaseId[caseId] > 1 || hasAmbiguousReportRow) {
      out[caseId] = {
        match_status: 'ambiguous',
        match_confidence: 'needs_review',
        match_candidate_count: candidateCountByCaseId[caseId]
      };
    }
  });

  candidates.forEach(function(candidate) {
    // A report row can be assigned only when it has one possible case, and the
    // case has one possible report row. Anything else needs manual review.
    if (candidateCountByCaseId[candidate.case_id] !== 1) return;
    if (candidateCountByReportKey[candidate.report_key] !== 1) return;

    candidate.match_status = 'matched';
    candidate.match_candidate_count = 1;
    out[candidate.case_id] = candidate;
  });

  return out;
}

function getCaseLeadMatchDates_(caseRow) {
  const dates = [
    toDateOnlyKey_(firstNonEmpty_(caseRow.created_at, caseRow.case_created_at)),
    toDateOnlyKey_(firstNonEmpty_(caseRow.opened_date, caseRow.case_opened_date))
  ].filter(Boolean);

  return dates.filter(function(value, index) {
    return dates.indexOf(value) === index;
  });
}

function toDateOnlyKey_(value) {
  const date = toDateOnlyMaybe_(value);
  if (!date) return '';

  return [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, '0'),
    String(date.getDate()).padStart(2, '0')
  ].join('-');
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

function normalizePhone_(value) {
  const digits = String(value || '').replace(/\D+/g, '');
  if (!digits) return '';
  return digits.length > 10 ? digits.slice(-10) : digits;
}

function formatFactCaseMasterColumns_() {
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.sheets.factCaseMaster);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  [
    'case_opened_date',
    'case_created_at',
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
    'lead_match_score',
    'lead_match_candidate_count'
  ].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('0.00');
    }
  });
}
