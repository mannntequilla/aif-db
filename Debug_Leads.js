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
