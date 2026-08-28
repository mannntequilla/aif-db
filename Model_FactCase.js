/**
 * Canonical case fact: exactly one row per MyCase case.
 *
 * This table holds one operational row per case. Case number and case name are
 * included for restricted management exception reporting; many-to-many
 * relationships remain in their respective bridge tables to prevent duplicate
 * case counts.
 */
function buildFactCase() {
  const cases = readSheetAsObjects_(CONFIG.sheets.rawCases);
  const clients = readSheetAsObjectsIfExists_(CONFIG.sheets.rawClients);
  const invoices = readSheetAsObjectsIfExists_(CONFIG.sheets.rawInvoices);
  const events = readSheetAsObjectsIfExists_(CONFIG.sheets.rawEvents);
  const staff = readSheetAsObjectsIfExists_(CONFIG.sheets.rawStaff);
  const customFields = readSheetAsObjectsIfExists_(CONFIG.sheets.rawCustomFields);
  const leadsReport = readSheetAsObjectsIfExists_(CONFIG.sheets.rawMyCaseLeadsReport);

  const clientsById = indexBy_(clients, 'id');
  const staffById = indexBy_(staff, 'id');
  const invoicesByCaseId = aggregateInvoicesByCaseId_(invoices);
  const eventCountByCaseId = aggregateEventCountByCaseId_(events);
  const leadMatches = buildLeadMatches_(cases, leadsReport, clientsById);
  const caseTypeCustomFieldId = getCustomFieldIdByName_(customFields, 'Case Type', 'case');

  const today = new Date();
  const rows = cases.map(function(caseRow) {
    const caseId = String(firstNonEmpty_(caseRow.id, caseRow.case_id)).trim();
    const openedDate = toDateOnlyMaybe_(firstNonEmpty_(caseRow.opened_date, caseRow.case_opened_date));
    const closedDate = toDateOnlyMaybe_(firstNonEmpty_(caseRow.closed_date, caseRow.case_closed_date));
    const eventCount = eventCountByCaseId[caseId] || 0;
    const staffing = resolveCaseStaffing_(caseRow, staffById);
    const leadMatch = leadMatches[caseId] || {};
    const financials = invoicesByCaseId[caseId] || {
      total_invoice_amount: 0,
      total_paid_so_far: 0,
      total_balance: 0
    };
    const status = String(firstNonEmpty_(caseRow.status, caseRow.case_status)).trim();
    const isClosed = normalizeText_(status) === 'closed' || Boolean(closedDate);
    const referenceDate = isClosed ? closedDate : today;
    const daysSinceOpen = daysBetweenDates_(openedDate, referenceDate);
    // Use the case record's own update timestamp as the operational recency
    // signal. A calendar event's scheduled start can be in the future and does
    // not reliably represent work performed on the case.
    const lastActivityDate = toDateOnlyMaybe_(firstNonEmpty_(
      caseRow.updated_at,
      caseRow.case_updated_at
    ));
    const daysSinceLastActivity = daysBetweenDates_(lastActivityDate, today);
    const linkedClient = resolveClientFromRef_(findPreferredCaseClientRef_(caseRow), clientsById) || {};
    const referralSource = firstNonEmpty_(leadMatch.referral_source);

    return {
      case_key: caseId,
      case_id: caseId,
      case_number: firstNonEmpty_(caseRow.case_number),
      case_name: firstNonEmpty_(caseRow.name, caseRow.case_name),

      case_type_key: normalizeDimensionKey_(getCaseCustomFieldValueById_(caseRow, caseTypeCustomFieldId)),
      practice_area_key: normalizeDimensionKey_(firstNonEmpty_(caseRow.practice_area, caseRow.practice_area_name)),
      current_case_stage_key: normalizeDimensionKey_(firstNonEmpty_(caseRow.case_stage, caseRow.stage)),
      case_status_key: normalizeDimensionKey_(status),
      office_key: String(firstNonEmpty_(caseRow.office_id, safeGet_(parseJsonMaybe_(caseRow.office), 'id', ''))).trim(),

      opened_date: openedDate,
      closed_date: closedDate,
      created_date: toDateOnlyMaybe_(firstNonEmpty_(caseRow.created_at, caseRow.case_created_at)),
      last_activity_date: lastActivityDate,

      primary_client_key: String(firstNonEmpty_(linkedClient.id, linkedClient.client_id)).trim(),
      primary_attorney_key: staffing.primary_attorney_id,
      lead_attorney_key: staffing.lead_attorney_id,
      lead_attorney_name: staffing.lead_attorney_name,
      originating_attorney_key: staffing.originating_attorney_id,
      referral_source_key: normalizeDimensionKey_(referralSource),
      referral_match_status: firstNonEmpty_(leadMatch.match_status, 'unmatched'),

      case_count: 1,
      is_open: isClosed ? 0 : 1,
      is_closed: isClosed ? 1 : 0,
      days_since_open: daysSinceOpen,
      days_to_close: isClosed ? daysSinceOpen : '',
      days_since_last_activity: daysSinceLastActivity,
      has_recent_activity_14d: daysSinceLastActivity !== '' && daysSinceLastActivity <= 14 ? 1 : 0,
      is_stale_14d: !isClosed && daysSinceLastActivity !== '' && daysSinceLastActivity > 14 ? 1 : 0,
      has_no_activity: lastActivityDate ? 0 : 1,
      event_count: eventCount,
      assigned_staff_count: staffing.assigned_staff_count,
      assigned_attorney_count: staffing.assigned_attorney_count,
      has_staff_assigned: staffing.assigned_staff_count > 0 ? 1 : 0,

      total_invoice_amount: financials.total_invoice_amount,
      total_paid_so_far: financials.total_paid_so_far,
      total_balance: financials.total_balance
    };
  });

  writeRowsToSheet_(CONFIG.sheets.factCase, rows);
  formatFactCaseColumns_();
}

function aggregateEventCountByCaseId_(events) {
  const out = {};

  events.forEach(function(eventRow) {
    const caseId = String(firstNonEmpty_(extractCaseIdFromEvent_(eventRow))).trim();
    if (!caseId) return;

    out[caseId] = (out[caseId] || 0) + 1;
  });

  return out;
}

function resolveCaseStaffing_(caseRow, staffById) {
  const assignments = parseJsonMaybe_(firstNonEmpty_(caseRow.staff, '[]'));
  const members = Array.isArray(assignments) ? assignments : [];
  const attorneyAssignments = members.filter(function(member) {
    const staffMember = staffById[String(firstNonEmpty_(member.id)).trim()] || {};
    return member.lead_lawyer === true || member.originating_lawyer === true ||
      normalizeText_(staffMember.title).indexOf('attorney') !== -1 ||
      normalizeText_(staffMember.role).indexOf('attorney') !== -1;
  });
  const leadAttorney = attorneyAssignments.find(function(member) {
    return member.lead_lawyer === true;
  });
  const originatingAttorney = attorneyAssignments.find(function(member) {
    return member.originating_lawyer === true;
  });
  const primaryAttorney = leadAttorney || originatingAttorney || attorneyAssignments[0] || {};
  const leadAttorneyId = String(firstNonEmpty_(leadAttorney && leadAttorney.id)).trim();
  const leadAttorneyStaff = staffById[leadAttorneyId] || {};

  return {
    lead_attorney_name: firstNonEmpty_(
      leadAttorneyStaff.full_name,
      buildFullName_(leadAttorneyStaff)
    ),
    primary_attorney_id: String(firstNonEmpty_(primaryAttorney.id)).trim(),
    lead_attorney_id: String(firstNonEmpty_(leadAttorney && leadAttorney.id)).trim(),
    originating_attorney_id: String(firstNonEmpty_(originatingAttorney && originatingAttorney.id)).trim(),
    assigned_staff_count: members.length,
    assigned_attorney_count: attorneyAssignments.length
  };
}

function daysBetweenDates_(startDate, endDate) {
  if (!startDate || !endDate) return '';
  const start = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  const end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  return Math.max(0, Math.floor((end - start) / 86400000));
}

function normalizeDimensionKey_(value) {
  return normalizeText_(value).replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

function formatFactCaseColumns_() {
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.sheets.factCase);
  if (!sheet || sheet.getLastRow() < 2) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  ['opened_date', 'closed_date', 'created_date', 'last_activity_date'].forEach(function(name) {
    const column = headers.indexOf(name) + 1;
    if (column > 0) sheet.getRange(2, column, sheet.getLastRow() - 1, 1).setNumberFormat('yyyy-mm-dd');
  });
  ['total_invoice_amount', 'total_paid_so_far', 'total_balance'].forEach(function(name) {
    const column = headers.indexOf(name) + 1;
    if (column > 0) sheet.getRange(2, column, sheet.getLastRow() - 1, 1).setNumberFormat('0.00');
  });
}
