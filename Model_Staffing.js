function buildCaseStaffTable() {
  const casesSheet = readSheetAsObjects_(CONFIG.sheets.rawCases);
  const staffSheet = readSheetAsObjects_(CONFIG.sheets.rawStaff);
  const staffById = {};

  staffSheet.forEach(function(s) {
    const id = String(s.id).trim();
    staffById[id] = s;
  });

  const output = [];

  casesSheet.forEach(function(c) {
    const caseId = c.id;
    const caseName = c.name;
    const caseStaff = c.staff;

    let assignedStaffNames = [];
    let assignedStaffIds = [];

    if (caseStaff) {
      const parsedStaff = parseJsonMaybe_(caseStaff);

      if (Array.isArray(parsedStaff)) {
        parsedStaff.forEach(function(member) {
          const staffId = String(member.id).trim();
          const staffMatch = staffById[staffId];

          const fullName = staffMatch
            ? [staffMatch.first_name, staffMatch.last_name].filter(Boolean).join(' ')
            : `ID ${staffId}`;

          assignedStaffNames.push(fullName);
          assignedStaffIds.push(staffId);
        });
      }
    }

    output.push({
      case_id: caseId,
      case_name: caseName,
      assigned_staff_names: assignedStaffNames.join(', '),
      assigned_staff_ids: assignedStaffIds.join(', '),
      has_staff_assigned: assignedStaffNames.length > 0 ? 'Yes' : 'No'
    });
  });

  writeRowsToSheet_('case_staff_summary', output);
}

/**
 * Current case-to-worker bridge. A case can appear more than once here, so
 * dashboard calculations must count distinct case_key after joining it to
 * fact_case; never sum fact_case.case_count through this bridge.
 */
function buildBridgeCaseStaff() {
  const cases = readSheetAsObjectsIfExists_(CONFIG.sheets.rawCases);
  const staff = readSheetAsObjectsIfExists_(CONFIG.sheets.rawStaff);
  const staffById = indexBy_(staff, 'id');
  const rows = [];

  cases.forEach(function(caseRow) {
    const caseKey = String(firstNonEmpty_(caseRow.id, caseRow.case_id)).trim();
    const assignments = parseJsonMaybe_(firstNonEmpty_(caseRow.staff, '[]'));
    if (!caseKey || !Array.isArray(assignments)) return;

    assignments.forEach(function(assignment) {
      const staffKey = String(firstNonEmpty_(assignment.id, assignment.staff_id)).trim();
      if (!staffKey) return;

      const staffRow = staffById[staffKey] || {};
      const staffName = firstNonEmpty_(
        staffRow.full_name,
        buildFullName_(staffRow)
      );
      const isLeadAttorney = assignment.lead_lawyer === true;
      const isOriginatingAttorney = assignment.originating_lawyer === true;
      const isAttorney = isLeadAttorney || isOriginatingAttorney ||
        normalizeText_(staffRow.title).indexOf('attorney') !== -1 ||
        normalizeText_(staffRow.role).indexOf('attorney') !== -1;

      rows.push({
        case_key: caseKey,
        staff_key: staffKey,
        // Display attribute for charts. Keep staff_key as the relationship key.
        staff_name: staffName || ('Unknown staff ' + staffKey),
        staff_title: firstNonEmpty_(staffRow.title),
        staff_active: firstNonEmpty_(staffRow.active),
        staffing_role: isLeadAttorney ? 'lead_attorney' :
          (isOriginatingAttorney ? 'originating_attorney' :
            (isAttorney ? 'attorney' : 'assigned_staff')),
        is_lead_attorney: isLeadAttorney ? 1 : 0,
        is_originating_attorney: isOriginatingAttorney ? 1 : 0,
        is_attorney: isAttorney ? 1 : 0
      });
    });
  });

  writeRowsToSheet_(CONFIG.sheets.bridgeCaseStaff, rows);
}

/**
 * Reporting-ready aggregation for the active workload chart. The fact is
 * joined to the staff bridge here, so dashboards do not have to implement a
 * many-to-many join or risk double-counting cases.
 */
function buildOpenCasesByParalegalReport() {
  const cases = readSheetAsObjectsIfExists_(CONFIG.sheets.rawCases);
  const assignments = readSheetAsObjectsIfExists_(CONFIG.sheets.bridgeCaseStaff);
  const caseByKey = indexBy_(cases, 'id');
  const grouped = {};

  assignments.forEach(function(assignment) {
    const staffTitle = String(firstNonEmpty_(assignment.staff_title)).trim();
    if (!isParalegal_(staffTitle)) return;

    const caseRow = caseByKey[String(firstNonEmpty_(assignment.case_key)).trim()] || {};
    if (!isOpenCase_(caseRow)) return;

    const staffKey = String(firstNonEmpty_(assignment.staff_key)).trim();
    if (!staffKey) return;
    if (!grouped[staffKey]) {
      grouped[staffKey] = {
        staff_key: staffKey,
        paralegal_name: firstNonEmpty_(assignment.staff_name, 'Unknown staff ' + staffKey),
        staff_title: staffTitle,
        open_case_keys: {}
      };
    }

    grouped[staffKey].open_case_keys[String(firstNonEmpty_(assignment.case_key))] = true;
  });

  const rows = Object.keys(grouped).map(function(staffKey) {
    const row = grouped[staffKey];
    return {
      staff_key: row.staff_key,
      paralegal_name: row.paralegal_name,
      staff_title: row.staff_title,
      open_case_count: Object.keys(row.open_case_keys).length
    };
  }).sort(function(a, b) {
    return b.open_case_count - a.open_case_count ||
      String(a.paralegal_name).localeCompare(String(b.paralegal_name));
  });

  writeRowsToSheet_(CONFIG.sheets.reportOpenCasesByParalegal, rows);
}

function isParalegal_(staffTitle) {
  return normalizeText_(staffTitle).indexOf('paralegal') !== -1;
}

function isOpenCase_(caseRow) {
  const status = normalizeText_(firstNonEmpty_(caseRow.status, caseRow.case_status));
  const closedDate = toDateOnlyMaybe_(firstNonEmpty_(caseRow.closed_date, caseRow.case_closed_date));
  return status !== 'closed' && !closedDate;
}
