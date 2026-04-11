function getProfitabilityExpenseActivityNames_() {
  return [
    'El abogado',
    'Facebook Ads',
    'Belgrano',
    'Spanish Smile'
  ];
}

function buildAllowedExpenseActivityLookup_() {
  const out = {};

  getProfitabilityExpenseActivityNames_().forEach(function(activityName) {
    out[normalizeText_(activityName)] = activityName;
  });

  return out;
}

function extractProfitabilityExpenseAmount_(expenseRow) {
  return toNumber_(
    firstNonEmpty_(
      expenseRow.amount,
      expenseRow.total_amount,
      expenseRow.value,
      expenseRow.expense_amount
    )
  );
}

function extractProfitabilityActivityName_(value) {
  return String(value || '').trim();
}

function formatCaseCreationMonth_(value) {
  const dateValue = toDateOnlyMaybe_(value);
  if (!dateValue) return '';

  return Utilities.formatDate(
    dateValue,
    Session.getScriptTimeZone(),
    'yyyy-MM'
  );
}

function aggregateAllowedExpensesByActivityName_(expenses) {
  const out = {};
  const allowedActivities = buildAllowedExpenseActivityLookup_();

  expenses.forEach(function(expenseRow) {
    const activityKey = normalizeText_(firstNonEmpty_(expenseRow.activity_name));

    if (!activityKey || !allowedActivities[activityKey]) return;

    if (!out[activityKey]) {
      out[activityKey] = {
        activity_name: allowedActivities[activityKey],
        expense_amount: 0
      };
    }

    out[activityKey].expense_amount += extractProfitabilityExpenseAmount_(expenseRow);
  });

  return out;
}

function buildFactCaseProfitability() {
  const caseMasterRows = readSheetAsObjectsIfExists_(CONFIG.sheets.factCaseMaster);
  const expenses = readSheetAsObjectsIfExists_(CONFIG.sheets.rawExpenses);
  const expensesByActivityName = aggregateAllowedExpensesByActivityName_(expenses);

  if (!caseMasterRows.length) {
    writeRowsToSheet_(CONFIG.sheets.factCaseProfitability, []);
    return;
  }

  const grouped = {};

  caseMasterRows.forEach(function(caseRow) {
    const activityName = extractProfitabilityActivityName_(caseRow.lead_referral_source);
    const activityKey = normalizeText_(activityName);
    const caseCreationDate = firstNonEmpty_(caseRow.case_opened_date);
    const caseCreationMonth = formatCaseCreationMonth_(caseCreationDate);

    if (!activityKey || !expensesByActivityName[activityKey]) return;

    const groupKey = [
      activityKey,
      formatDateOnlyForSheet_(caseCreationDate),
      caseCreationMonth
    ].join('|');

    if (!grouped[groupKey]) {
      grouped[groupKey] = {
        activity_name: expensesByActivityName[activityKey].activity_name,
        case_creation_date: formatDateOnlyForSheet_(caseCreationDate),
        case_creation_month: caseCreationMonth,
        case_count: 0,
        retainer_amount: 0,
        expense_amount: 0
      };
    }

    grouped[groupKey].case_count += 1;
    grouped[groupKey].retainer_amount += toNumber_(caseRow.retainer);
  });

  const rows = Object.keys(grouped)
    .sort()
    .map(function(groupKey) {
      const row = grouped[groupKey];
      row.expense_amount = expensesByActivityName[normalizeText_(row.activity_name)].expense_amount;
      row.net_profit = row.retainer_amount - row.expense_amount;
      row.roi = row.expense_amount ? row.net_profit / row.expense_amount : '';
      row.roas = row.expense_amount ? row.retainer_amount / row.expense_amount : '';
      return row;
    });

  writeRowsToSheet_(CONFIG.sheets.factCaseProfitability, rows);
  formatFactCaseProfitabilityColumns_();
}

function formatFactCaseProfitabilityColumns_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.factCaseProfitability);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  [
    'case_count',
    'retainer_amount',
    'expense_amount',
    'net_profit',
    'roi',
    'roas'
  ].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('0.00');
    }
  });
}
