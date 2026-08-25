/**
 * Reusable workload mart with one row per case and assigned worker. It brings
 * current case attributes into the assignment grain, so charts can filter and
 * group directly from this sheet without creating a new table per chart.
 */
function buildCaseWorkloadByStaff() {
  const cases = readSheetAsObjectsIfExists_(CONFIG.sheets.rawCases);
  const staff = readSheetAsObjectsIfExists_(CONFIG.sheets.rawStaff);
  const factCases = readSheetAsObjectsIfExists_(CONFIG.sheets.factCase);
  const operationsStaff = readSheetAsObjectsIfExists_(CONFIG.sheets.operationsStaffRoster);
  const staffById = indexBy_(staff, 'id');
  const factCaseByKey = indexBy_(factCases, 'case_key');
  const operationsStaffByKey = indexBy_(operationsStaff, 'staff_key');
  const rows = [];
  const seenAssignments = {};

  migrateBridgeCaseStaffSheet_();

  cases.forEach(function(caseRow) {
    const caseKey = String(firstNonEmpty_(caseRow.id, caseRow.case_id)).trim();
    const assignments = parseJsonMaybe_(firstNonEmpty_(caseRow.staff, '[]'));
    if (!caseKey || !Array.isArray(assignments)) return;

    assignments.forEach(function(assignment) {
      const staffKey = String(firstNonEmpty_(assignment.id, assignment.staff_id)).trim();
      if (!staffKey) return;

      const assignmentKey = [caseKey, staffKey].join('|');
      if (seenAssignments[assignmentKey]) return;
      seenAssignments[assignmentKey] = true;

      const staffRow = staffById[staffKey] || {};
      const operationsStaffRow = operationsStaffByKey[staffKey] || {};
      const factCase = factCaseByKey[caseKey] || {};
      const staffName = firstNonEmpty_(
        staffRow.full_name,
        buildFullName_(staffRow)
      );
      const isLeadAttorney = assignment.lead_lawyer === true;
      const isOriginatingAttorney = assignment.originating_lawyer === true;
      // In this MyCase setup, lead_lawyer denotes the operational paralegal
      // responsible for the case; it is not a reliable attorney classification.
      const isAttorney = !isLeadAttorney && (
        normalizeText_(staffRow.title).indexOf('attorney') !== -1 ||
        normalizeText_(staffRow.role).indexOf('attorney') !== -1
      );

      rows.push({
        case_key: caseKey,
        case_id: firstNonEmpty_(factCase.case_id, caseKey),
        staff_key: staffKey,
        staff_name: staffName || ('Unknown staff ' + staffKey),
        staff_title: firstNonEmpty_(staffRow.title),
        staff_active: firstNonEmpty_(staffRow.active),
        operational_role: firstNonEmpty_(operationsStaffRow.operational_role, 'unclassified'),
        is_paralegal: firstNonEmpty_(operationsStaffRow.is_paralegal, 'No'),
        is_case_owner: firstNonEmpty_(operationsStaffRow.is_case_owner, 'No'),
        staffing_role: isLeadAttorney ? 'lead_paralegal' :
          (isOriginatingAttorney ? 'originating_attorney' :
            (isAttorney ? 'attorney' : 'assigned_staff')),
        is_lead_attorney: isLeadAttorney ? 1 : 0,
        is_originating_attorney: isOriginatingAttorney ? 1 : 0,
        is_attorney: isAttorney ? 1 : 0,

        case_type_key: firstNonEmpty_(factCase.case_type_key),
        practice_area_key: firstNonEmpty_(factCase.practice_area_key),
        current_case_stage_key: firstNonEmpty_(factCase.current_case_stage_key),
        case_status_key: firstNonEmpty_(factCase.case_status_key),
        opened_date: firstNonEmpty_(factCase.opened_date),
        days_since_open: firstNonEmpty_(factCase.days_since_open),
        days_since_last_activity: firstNonEmpty_(factCase.days_since_last_activity),
        is_open: firstNonEmpty_(factCase.is_open),
        is_stale_14d: firstNonEmpty_(factCase.is_stale_14d)
      });
    });
  });

  writeRowsToSheet_(CONFIG.sheets.caseWorkloadByStaff, rows);
}

// Renames the previous technical sheet on its first rebuild, retaining any
// existing charts or formulas that reference the sheet by name.
function migrateBridgeCaseStaffSheet_() {
  const spreadsheet = getSpreadsheet_();
  const legacySheet = spreadsheet.getSheetByName('bridge_case_staff');
  const workloadSheet = spreadsheet.getSheetByName(CONFIG.sheets.caseWorkloadByStaff);

  if (legacySheet && !workloadSheet) {
    legacySheet.setName(CONFIG.sheets.caseWorkloadByStaff);
  }
}
