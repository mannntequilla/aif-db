/**
 * Current-state foundations for the Paralegal Operations Overview.
 * Historical stage movement, deadlines and individual performance are
 * intentionally out of scope until their source data is modeled.
 */
function buildOperationsOverviewData_() {
  buildOperationsStaffRoster_();
  buildCaseWorkloadByStaff();
  buildFactCaseAttention_();
}

/**
 * Creates an editable roster. Manual operational classifications are retained
 * across refreshes and are never inferred from job title or staff type.
 */
function buildOperationsStaffRoster_() {
  const staff = readSheetAsObjectsIfExists_(CONFIG.sheets.rawStaff);
  const existing = readSheetAsObjectsIfExists_(CONFIG.sheets.operationsStaffRoster);
  const existingByStaffKey = indexBy_(existing, 'staff_key');

  const rows = staff.map(function(staffRow) {
    const staffKey = String(firstNonEmpty_(staffRow.id)).trim();
    const current = existingByStaffKey[staffKey] || {};

    return {
      staff_key: staffKey,
      staff_name: firstNonEmpty_(staffRow.full_name, buildFullName_(staffRow)),
      source_title: firstNonEmpty_(staffRow.title),
      source_type: firstNonEmpty_(staffRow.type),
      staff_active: firstNonEmpty_(staffRow.active),
      operational_role: firstNonEmpty_(current.operational_role, 'unclassified'),
      is_paralegal: firstNonEmpty_(current.is_paralegal, 'No'),
      is_case_owner: firstNonEmpty_(current.is_case_owner, 'No'),
      supervisor: firstNonEmpty_(current.supervisor),
      notes: firstNonEmpty_(current.notes)
    };
  });

  writeRowsToSheet_(CONFIG.sheets.operationsStaffRoster, rows);
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.sheets.operationsStaffRoster);
  if (sheet) sheet.setFrozenRows(1);
}

/**
 * One row per open case with transparent current-state exception flags. A
 * case may have several reasons and is not attributed to an individual.
 */
function buildFactCaseAttention_() {
  const factCases = readSheetAsObjectsIfExists_(CONFIG.sheets.factCase);
  const rawCases = readSheetAsObjectsIfExists_(CONFIG.sheets.rawCases);
  const caseMasters = readSheetAsObjectsIfExists_(CONFIG.sheets.factCaseMaster);
  const assignments = readSheetAsObjectsIfExists_(CONFIG.sheets.caseWorkloadByStaff);
  const rawCaseById = indexBy_(rawCases, 'id');
  const caseMasterById = indexBy_(caseMasters, 'case_id');
  const staffNamesByCaseKey = {};

  assignments.forEach(function(row) {
    const caseKey = String(firstNonEmpty_(row.case_key)).trim();
    if (!caseKey) return;
    if (!staffNamesByCaseKey[caseKey]) staffNamesByCaseKey[caseKey] = [];
    if (row.staff_name) staffNamesByCaseKey[caseKey].push(row.staff_name);
  });

  const rows = factCases.filter(function(row) {
    return String(firstNonEmpty_(row.is_open)) === '1';
  }).map(function(caseRow) {
    const caseKey = String(firstNonEmpty_(caseRow.case_key, caseRow.case_id)).trim();
    const rawCase = rawCaseById[caseKey] || {};
    const caseMaster = caseMasterById[caseKey] || {};
    const flags = getCurrentCaseAttentionFlags_(caseRow);

    return {
      case_key: caseKey,
      case_id: firstNonEmpty_(caseRow.case_id, caseKey),
      case_number: firstNonEmpty_(rawCase.case_number),
      case_name: firstNonEmpty_(rawCase.name, caseMaster.case_name),
      assigned_staff: uniqueStrings_(staffNamesByCaseKey[caseKey] || []).join(', '),
      practice_area: firstNonEmpty_(caseMaster.practice_area, caseRow.practice_area_key),
      case_type: firstNonEmpty_(caseMaster.case_type, caseRow.case_type_key),
      current_case_stage: firstNonEmpty_(caseMaster.case_stage, caseRow.current_case_stage_key),
      case_status: firstNonEmpty_(caseMaster.case_status, caseRow.case_status_key),
      opened_date: firstNonEmpty_(caseRow.opened_date),
      days_since_open: firstNonEmpty_(caseRow.days_since_open),
      last_activity_date: firstNonEmpty_(caseRow.last_activity_date),
      days_since_last_activity: firstNonEmpty_(caseRow.days_since_last_activity),
      attention_severity: flags.severity,
      attention_category: flags.categories.join(' | '),
      attention_reason: flags.reasons.join(' | '),
      attention_rule_count: flags.reasons.length,
      is_requiring_attention: flags.reasons.length ? 1 : 0
    };
  }).filter(function(row) {
    return row.is_requiring_attention === 1;
  });

  rows.sort(function(a, b) {
    return attentionSeverityRank_(a.attention_severity) - attentionSeverityRank_(b.attention_severity) ||
      Number(b.days_since_last_activity || 0) - Number(a.days_since_last_activity || 0);
  });
  writeRowsToSheet_(CONFIG.sheets.caseAttention, rows);
}

function getCurrentCaseAttentionFlags_(caseRow) {
  const reasons = [];
  const categories = [];
  let severity = 'none';
  const daysSinceOpen = Number(firstNonEmpty_(caseRow.days_since_open, 0));

  function addFlag(category, reason, flagSeverity) {
    categories.push(category);
    reasons.push(reason);
    if (attentionSeverityRank_(flagSeverity) < attentionSeverityRank_(severity)) severity = flagSeverity;
  }

  if (String(firstNonEmpty_(caseRow.has_staff_assigned)) !== '1') {
    addFlag('Ownership', 'Open case has no staff assigned', 'critical');
  }
  if (!firstNonEmpty_(caseRow.current_case_stage_key)) {
    addFlag('Data quality', 'Current case stage is missing', 'high');
  }
  if (String(firstNonEmpty_(caseRow.has_no_activity)) === '1' && daysSinceOpen > 14) {
    addFlag('Activity', 'No recorded event activity more than 14 days after opening', 'high');
  } else if (String(firstNonEmpty_(caseRow.is_stale_14d)) === '1') {
    addFlag('Activity', 'No recorded event activity in the last 14 days', 'medium');
  }
  if (!firstNonEmpty_(caseRow.practice_area_key) || !firstNonEmpty_(caseRow.case_type_key)) {
    addFlag('Data quality', 'Practice area or case type is missing', 'low');
  }

  return { severity: severity, categories: uniqueStrings_(categories), reasons: reasons };
}

function attentionSeverityRank_(value) {
  const ranks = { critical: 1, high: 2, medium: 3, low: 4, none: 5 };
  return ranks[String(value || 'none').toLowerCase()] || 5;
}

function uniqueStrings_(values) {
  const seen = {};
  return values.filter(function(value) {
    const text = String(value || '').trim();
    if (!text || seen[text]) return false;
    seen[text] = true;
    return true;
  });
}
