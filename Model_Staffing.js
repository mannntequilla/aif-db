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
