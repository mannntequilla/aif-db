function debugRawLeadsHeaders() {
  const leads = readSheetAsObjects_(CONFIG.sheets.rawLeads);
  if (!leads.length) {
    Logger.log('No leads found');
    return;
  }

  Logger.log(JSON.stringify(Object.keys(leads[0])));
}

function debugLeadClientUuidOverlap() {
  const leads = readSheetAsObjects_(CONFIG.sheets.rawLeads);
  const clients = readSheetAsObjects_(CONFIG.sheets.rawClients);

  const leadUuids = new Set(
    leads
      .map(function(r) { return String(firstNonEmpty_(r.uuid)).trim(); })
      .filter(Boolean)
  );

  const matchingClients = clients.filter(function(client) {
    const uuid = String(firstNonEmpty_(client.uuid)).trim();
    return uuid && leadUuids.has(uuid);
  });

  Logger.log('Lead UUID count: ' + leadUuids.size);
  Logger.log('Matching clients by UUID: ' + matchingClients.length);

  matchingClients.slice(0, 20).forEach(function(client) {
    Logger.log(JSON.stringify({
      client_id: client.id,
      client_uuid: client.uuid,
      email: client.email,
      first_name: client.first_name,
      last_name: client.last_name
    }));
  });
}

function debugHeadersRawMyCaseLeadsReport() {
  const sheetName = CONFIG.sheets.rawMyCaseLeadsReport;

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('No existe la hoja: ' + sheetName);
  }

  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  Logger.log('HEADERS:');
  headers.forEach(function(h, i) {
    Logger.log((i + 1) + ': [' + h + ']');
  });

  return headers;
}

function getAllSpreadsheetHeaders() {
  const ss = getSpreadsheet_();
  const sheets = ss.getSheets();

  const result = sheets.map(function(sheet) {
    const lastColumn = sheet.getLastColumn();

    if (lastColumn === 0) {
      return {
        sheetName: sheet.getName(),
        headers: []
      };
    }

    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    return {
      sheetName: sheet.getName(),
      headers: headers
    };
  });

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function debugConvertedLeadClassification() {
  const rows = readSheetAsObjectsIfExists_(CONFIG.sheets.rawMyCaseLeadsReport);

  if (!rows || !rows.length) {
    writeRowsToSheet_('debug_converted_leads', []);
    Logger.log('No rows found in rawMyCaseLeadsReport');
    return;
  }

  const output = [];

  rows.forEach(function(row, index) {
    const rawStatus = firstNonEmpty_(row['Lead status']);
    const status = rawStatus ? String(rawStatus).trim() : '';
    const normalizedStatus = normalizeLeadStatus_(status);

    const rawConversionDate = firstNonEmpty_(row['Conversion date']);
    const parsedConversionDate = toDateOnlyMaybe_(rawConversionDate);

    const rawDateAdded = firstNonEmpty_(row['Date added']);
    const parsedDateAdded = toDateOnlyMaybe_(rawDateAdded);

    const classifiedStage = classifyLeadFunnelStage_(status, parsedConversionDate);

    const isConvertedStatus =
      normalizedStatus === 'contract' ||
      normalizedStatus === 'detainee visitation';

    let reason = '';

    if (isConvertedStatus) {
      if (parsedConversionDate) {
        reason = 'Converted status + valid parsed conversion date';
      } else {
        reason = 'Converted status but missing/invalid parsed conversion date';
      }

      output.push({
        row_number: index + 2,
        lead_name: firstNonEmpty_(row['Lead name']),
        lead_id: firstNonEmpty_(row['Lead ID']) || firstNonEmpty_(row['Id']) || '',
        raw_status: rawStatus,
        normalized_status: normalizedStatus,
        raw_conversion_date: rawConversionDate,
        parsed_conversion_date: parsedConversionDate,
        raw_date_added: rawDateAdded,
        parsed_date_added: parsedDateAdded,
        classified_stage: classifiedStage,
        reason: reason
      });
    }
  });

  writeRowsToSheet_('debug_converted_leads', output);
  Logger.log('Debug rows written: ' + output.length);
}

function showAccessToken() {
  const token = getAccessToken_();
  Logger.log(token);
}

function profileExpensesRaw_() {
  const expenses = readSheetAsObjectsIfExists_(CONFIG.sheets.rawExpenses);

  if (!expenses.length) {
    Logger.log('No expenses found in raw_expenses');
    writeRowsToSheet_('debug_expenses_profile', []);
    return;
  }

  Logger.log('Total expenses: ' + expenses.length);
  Logger.log('Expense headers: ' + JSON.stringify(Object.keys(expenses[0])));
  Logger.log('First expense sample: ' + JSON.stringify(expenses[0], null, 2));

  const output = expenses.slice(0, 200).map(function(expense) {
    return {
      expense_id: firstNonEmpty_(expense.id),
      case_id: firstNonEmpty_(expense.case_id),
      amount: firstNonEmpty_(expense.amount, expense.value, expense.total_amount),
      description: firstNonEmpty_(expense.description, expense.name, expense.title),
      expense_type: firstNonEmpty_(expense.expense_type, expense.type, expense.category),
      created_at: firstNonEmpty_(expense.created_at),
      updated_at: firstNonEmpty_(expense.updated_at),
      raw_case: asJson_(expense.case),
      raw_client: asJson_(expense.client)
    };
  });

  writeRowsToSheet_('debug_expenses_profile', output);
}

function debugCaseCoreFieldsTable() {
  const cases = readSheetAsObjectsIfExists_(CONFIG.sheets.rawCases);
  const customFields = readSheetAsObjectsIfExists_(CONFIG.sheets.rawCustomFields);
  const roles = readSheetAsObjectsIfExists_(CONFIG.sheets.rawRoles);
  const staff = readSheetAsObjectsIfExists_(CONFIG.sheets.rawStaff);
  const staffById = indexBy_(staff, 'id');
  const leadAttorneyByCaseId = buildLeadAttorneyByCaseId_(roles, staffById);
  const caseTypeField = findCustomFieldByName_(customFields, 'Case Type', 'case');
  const alienNumberField = findCustomFieldByName_(customFields, 'Alien Number', 'case');

  const rows = cases.map(function(caseRow) {
    const caseId = String(firstNonEmpty_(caseRow.id, caseRow.case_id)).trim();

    return {
      case_id: caseId,
      case_name: firstNonEmpty_(caseRow.name, caseRow.case_name),
      open_date: toDateOnlyMaybe_(firstNonEmpty_(caseRow.opened_date, caseRow.case_opened_date)),
      close_date: toDateOnlyMaybe_(firstNonEmpty_(caseRow.closed_date, caseRow.close_date, caseRow.closed_at, caseRow.close_at)),
      stage: firstNonEmpty_(caseRow.case_stage, caseRow.stage),
      case_type: resolveCaseCustomFieldDisplayValue_(caseRow, caseTypeField),
      alien_number: resolveCaseCustomFieldDisplayValue_(caseRow, alienNumberField),
      lead_attorney: firstNonEmpty_(leadAttorneyByCaseId[caseId], '')
    };
  });

  writeRowsToSheet_('debug_case_core_fields', rows);
  formatDebugCaseCoreFieldsColumns_();
}

function findCustomFieldByName_(customFields, fieldName, parentType) {
  if (!customFields || !customFields.length) return null;

  const normalizedFieldName = normalizeText_(fieldName);
  const normalizedParentType = normalizeText_(parentType || '');

  return customFields.find(function(customField) {
    const currentName = normalizeText_(customField.name);
    const currentParentType = normalizeText_(customField.parent_type);

    if (currentName !== normalizedFieldName) return false;
    if (normalizedParentType && currentParentType !== normalizedParentType) return false;

    return true;
  }) || null;
}

function resolveCaseCustomFieldDisplayValue_(caseRow, customFieldRow) {
  if (!customFieldRow) return '';

  const rawValue = getCaseCustomFieldValueById_(caseRow, firstNonEmpty_(customFieldRow.id));
  if (rawValue === '' || rawValue === null || rawValue === undefined) return '';

  const optionLabelById = buildCustomFieldOptionLabelById_(customFieldRow);
  const parsedValue = parseJsonMaybe_(rawValue);

  if (Array.isArray(parsedValue)) {
    return parsedValue.map(function(item) {
      return resolveCustomFieldSingleValue_(item, optionLabelById);
    }).filter(Boolean).join(', ');
  }

  if (parsedValue && typeof parsedValue === 'object') {
    return resolveCustomFieldSingleValue_(parsedValue, optionLabelById);
  }

  return resolveCustomFieldSingleValue_(rawValue, optionLabelById);
}

function resolveCustomFieldSingleValue_(value, optionLabelById) {
  const isObjectValue = value && typeof value === 'object';
  const optionId = String(
    firstNonEmpty_(
      isObjectValue ? safeGet_(value, 'id', '') : '',
      isObjectValue ? safeGet_(value, 'value', '') : '',
      value
    )
  ).trim();

  if (!optionId) return '';
  return firstNonEmpty_(optionLabelById[optionId], optionId);
}

function buildCustomFieldOptionLabelById_(customFieldRow) {
  const out = {};
  const options = firstNonEmpty_(
    parseJsonMaybe_(customFieldRow.options),
    parseJsonMaybe_(customFieldRow.values),
    parseJsonMaybe_(customFieldRow.choices),
    []
  );

  if (!Array.isArray(options)) return out;

  options.forEach(function(optionRow) {
    const optionId = String(firstNonEmpty_(optionRow.id, optionRow.value)).trim();
    const optionLabel = String(firstNonEmpty_(optionRow.name, optionRow.label, optionRow.value)).trim();

    if (!optionId) return;
    out[optionId] = optionLabel;
  });

  return out;
}

function buildLeadAttorneyByCaseId_(roles, staffById) {
  const out = {};

  roles.forEach(function(roleRow) {
    const roleBlob = JSON.stringify(roleRow || {}).toLowerCase();
    if (roleBlob.indexOf('lead attorney') === -1) return;

    const caseId = String(
      firstNonEmpty_(
        roleRow.case_id,
        safeGet_(parseJsonMaybe_(roleRow.case), 'id', '')
      )
    ).trim();
    if (!caseId || out[caseId]) return;

    out[caseId] = firstNonEmpty_(
      resolveRoleStaffName_(roleRow, staffById),
      String(firstNonEmpty_(roleRow.name, roleRow.title)).trim()
    );
  });

  return out;
}

function resolveRoleStaffName_(roleRow, staffById) {
  const possibleRefs = [
    roleRow.staff,
    roleRow.user,
    roleRow.person,
    roleRow.assignee,
    roleRow.member
  ];

  for (var i = 0; i < possibleRefs.length; i++) {
    const parsedRef = parseJsonMaybe_(possibleRefs[i]) || possibleRefs[i];
    const staffId = String(firstNonEmpty_(safeGet_(parsedRef, 'id', ''))).trim();

    if (staffId && staffById[staffId]) {
      return [staffById[staffId].first_name, staffById[staffId].last_name].filter(Boolean).join(' ');
    }

    const directName = String(
      firstNonEmpty_(
        safeGet_(parsedRef, 'full_name', ''),
        [
          safeGet_(parsedRef, 'first_name', ''),
          safeGet_(parsedRef, 'last_name', '')
        ].filter(Boolean).join(' ')
      )
    ).trim();

    if (directName) return directName;
  }

  return '';
}

function formatDebugCaseCoreFieldsColumns_() {
  const sheet = getSpreadsheet_().getSheetByName('debug_case_core_fields');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  ['open_date', 'close_date'].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
  });
}
