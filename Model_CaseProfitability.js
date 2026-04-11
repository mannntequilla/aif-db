function getProfitabilityGroupDefinitions_() {
  return [
    {
      profit_group: 'El Abogado.com',
      referral_sources: ['El Abogado.com'],
      expense_activity_names: ['El abogado']
    },
    {
      profit_group: 'Spanish Smile',
      referral_sources: ['Spanish Smile'],
      expense_activity_names: ['Spanish Smile', 'Facebook Ads']
    },
    {
      profit_group: 'Belgrano',
      referral_sources: [],
      expense_activity_names: ['Belgrano']
    },
    {
      profit_group: 'Google Ads',
      referral_sources: [],
      expense_activity_names: ['Google Ads']
    }
  ];
}

function buildProfitabilityReferralSourceLookup_() {
  const out = {};

  getProfitabilityGroupDefinitions_().forEach(function(definition) {
    definition.referral_sources.forEach(function(referralSource) {
      out[normalizeText_(referralSource)] = definition.profit_group;
    });
  });

  return out;
}

function buildProfitabilityExpenseActivityLookup_() {
  const out = {};

  getProfitabilityGroupDefinitions_().forEach(function(definition) {
    definition.expense_activity_names.forEach(function(activityName) {
      out[normalizeText_(activityName)] = definition.profit_group;
    });
  });

  return out;
}

function formatProfitabilityMetricMonth_(value) {
  const dateValue = toDateOnlyMaybe_(value);
  if (!dateValue) return '';

  return Utilities.formatDate(
    dateValue,
    Session.getScriptTimeZone(),
    'yyyy-MM'
  );
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

function extractProfitabilityExpenseDate_(expenseRow) {
  return firstNonEmpty_(
    expenseRow.expense_date,
    expenseRow.date,
    expenseRow.incurred_on,
    expenseRow.created_at
  );
}

function initializeProfitabilityGroupRow_(profitGroup, metricDate, metricMonth) {
  return {
    profit_group: profitGroup,
    metric_date: metricDate,
    metric_month: metricMonth,
    retainer_amount: 0,
    expense_amount: 0,
    net_profit: 0,
    roi: '',
    roas: '',
    case_count: 0
  };
}

function aggregateProfitabilityRevenue_(caseMasterRows) {
  const grouped = {};
  const referralSourceLookup = buildProfitabilityReferralSourceLookup_();

  caseMasterRows.forEach(function(caseRow) {
    const referralSourceKey = normalizeText_(caseRow.lead_referral_source);
    const profitGroup = referralSourceLookup[referralSourceKey];
    const metricDate = formatDateOnlyForSheet_(firstNonEmpty_(caseRow.case_opened_date));
    const metricMonth = formatProfitabilityMetricMonth_(metricDate);

    if (!profitGroup || !metricDate) return;

    const groupKey = [profitGroup, metricDate, metricMonth].join('|');

    if (!grouped[groupKey]) {
      grouped[groupKey] = initializeProfitabilityGroupRow_(profitGroup, metricDate, metricMonth);
    }

    grouped[groupKey].retainer_amount += toNumber_(caseRow.retainer);
    grouped[groupKey].case_count += 1;
  });

  return grouped;
}

function aggregateProfitabilityExpenses_(expenses) {
  const grouped = {};
  const expenseActivityLookup = buildProfitabilityExpenseActivityLookup_();

  expenses.forEach(function(expenseRow) {
    const activityNameKey = normalizeText_(firstNonEmpty_(expenseRow.activity_name));
    const profitGroup = expenseActivityLookup[activityNameKey];
    const metricDate = formatDateOnlyForSheet_(extractProfitabilityExpenseDate_(expenseRow));
    const metricMonth = formatProfitabilityMetricMonth_(metricDate);

    if (!profitGroup || !metricDate) return;

    const groupKey = [profitGroup, metricDate, metricMonth].join('|');

    if (!grouped[groupKey]) {
      grouped[groupKey] = initializeProfitabilityGroupRow_(profitGroup, metricDate, metricMonth);
    }

    grouped[groupKey].expense_amount += extractProfitabilityExpenseAmount_(expenseRow);
  });

  return grouped;
}

function mergeProfitabilityGroups_(revenueGroups, expenseGroups) {
  const merged = {};

  Object.keys(revenueGroups).forEach(function(groupKey) {
    merged[groupKey] = Object.assign({}, revenueGroups[groupKey]);
  });

  Object.keys(expenseGroups).forEach(function(groupKey) {
    if (!merged[groupKey]) {
      merged[groupKey] = Object.assign({}, expenseGroups[groupKey]);
      return;
    }

    merged[groupKey].expense_amount += expenseGroups[groupKey].expense_amount;
  });

  return merged;
}

function buildFactCaseProfitability() {
  const caseMasterRows = readSheetAsObjectsIfExists_(CONFIG.sheets.factCaseMaster);
  const expenses = readSheetAsObjectsIfExists_(CONFIG.sheets.rawExpenses);

  const revenueGroups = aggregateProfitabilityRevenue_(caseMasterRows);
  const expenseGroups = aggregateProfitabilityExpenses_(expenses);
  const grouped = mergeProfitabilityGroups_(revenueGroups, expenseGroups);

  const rows = Object.keys(grouped)
    .sort()
    .map(function(groupKey) {
      const row = grouped[groupKey];
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
    'retainer_amount',
    'expense_amount',
    'net_profit',
    'roi',
    'roas',
    'case_count'
  ].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('0.00');
    }
  });
}
