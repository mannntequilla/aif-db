function getProfitabilityExpenseActivityNames_() {
  return [
    'El abogado',
    'Facebook Ads',
    'Belgrano',
    'Spanish Smile'
  ];
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

function extractProfitabilityExpenseCaseId_(expenseRow) {
  const caseObj = parseJsonMaybe_(expenseRow.case);

  return firstNonEmpty_(
    expenseRow.case_id,
    caseObj && caseObj.id
  );
}

function buildAllowedExpenseActivityLookup_() {
  const out = {};

  getProfitabilityExpenseActivityNames_().forEach(function(activityName) {
    out[normalizeText_(activityName)] = activityName;
  });

  return out;
}

function aggregateAllowedExpensesByCaseId_(expenses) {
  const out = {};
  const allowedActivities = buildAllowedExpenseActivityLookup_();

  expenses.forEach(function(expenseRow) {
    const caseId = String(extractProfitabilityExpenseCaseId_(expenseRow) || '');
    const activityName = normalizeText_(firstNonEmpty_(expenseRow.activity_name));

    if (!caseId || !activityName || !allowedActivities[activityName]) return;

    if (!out[caseId]) {
      out[caseId] = {
        filtered_expense_amount: 0,
        filtered_expense_activity_names: {}
      };
    }

    out[caseId].filtered_expense_amount += extractProfitabilityExpenseAmount_(expenseRow);
    out[caseId].filtered_expense_activity_names[allowedActivities[activityName]] = true;
  });

  return out;
}

function buildFactCaseProfitability() {
  const caseMasterRows = readSheetAsObjectsIfExists_(CONFIG.sheets.factCaseMaster);
  const expenses = readSheetAsObjectsIfExists_(CONFIG.sheets.rawExpenses);
  const expensesByCaseId = aggregateAllowedExpensesByCaseId_(expenses);

  if (!caseMasterRows.length) {
    writeRowsToSheet_(CONFIG.sheets.factCaseProfitability, []);
    return;
  }

  const rows = caseMasterRows.map(function(caseRow) {
    const caseId = String(firstNonEmpty_(caseRow.case_id) || '');
    const expenseSummary = expensesByCaseId[caseId] || {
      filtered_expense_amount: 0,
      filtered_expense_activity_names: {}
    };
    const retainerAmount = toNumber_(caseRow.retainer);
    const expenseAmount = toNumber_(expenseSummary.filtered_expense_amount);
    const netProfit = retainerAmount - expenseAmount;
    const roi = expenseAmount ? netProfit / expenseAmount : '';
    const roas = expenseAmount ? retainerAmount / expenseAmount : '';

    return {
      case_id: caseId,
      case_name: firstNonEmpty_(caseRow.case_name),
      client_full_name: firstNonEmpty_(caseRow.client_full_name),
      case_status: firstNonEmpty_(caseRow.case_status),
      case_stage: firstNonEmpty_(caseRow.case_stage),
      lead_referral_source: firstNonEmpty_(caseRow.lead_referral_source),
      retainer_amount: retainerAmount,
      filtered_expense_amount: expenseAmount,
      filtered_expense_activity_names: Object.keys(expenseSummary.filtered_expense_activity_names).join(', '),
      net_profit: netProfit,
      roi: roi,
      roas: roas
    };
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
    'filtered_expense_amount',
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
