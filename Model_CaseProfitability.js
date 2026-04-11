function extractCaseProfitabilityExpenseAmount_(expenseRow) {
  return toNumber_(
    firstNonEmpty_(
      expenseRow.cost,
      expenseRow.amount,
      expenseRow.total_amount,
      expenseRow.value,
      expenseRow.expense_amount
    )
  );
}

function extractCaseProfitabilityEntryDate_(expenseRow) {
  return formatDateOnlyForSheet_(firstNonEmpty_(expenseRow.entry_date));
}

function formatCaseProfitabilityEntryMonth_(value) {
  const dateValue = toDateOnlyMaybe_(value);
  if (!dateValue) return '';

  return Utilities.formatDate(
    dateValue,
    Session.getScriptTimeZone(),
    'yyyy-MM'
  );
}

function getAllowedCaseProfitabilityActivityNames_() {
  return [
    'Belgrano',
    'Spanish Smile',
    'El abogado',
    'Facebook Ads'
  ];
}

function isAllowedCaseProfitabilityActivityName_(activityName) {
  const normalizedActivityName = normalizeText_(activityName);

  return getAllowedCaseProfitabilityActivityNames_().some(function(allowedName) {
    return normalizedActivityName === normalizeText_(allowedName);
  });
}

function getLinkedReferralSourceByActivityName_(activityName) {
  const normalizedActivityName = normalizeText_(activityName);

  if (normalizedActivityName === normalizeText_('Belgrano')) {
    return 'Google Ads';
  }

  if (normalizedActivityName === normalizeText_('Spanish Smile')) {
    return 'Spanish Smile';
  }

  if (normalizedActivityName === normalizeText_('El abogado')) {
    return 'El Abogado.com';
  }

  if (normalizedActivityName === normalizeText_('Facebook Ads')) {
    return 'Spanish Smile';
  }

  return '';
}

function aggregateRetainersByReferralSourceAndMonth_(caseMasterRows) {
  const out = {};

  caseMasterRows.forEach(function(caseRow) {
    const referralSource = String(firstNonEmpty_(caseRow.lead_referral_source)).trim();
    const referralSourceKey = normalizeText_(referralSource);
    const metricMonth = formatCaseProfitabilityEntryMonth_(firstNonEmpty_(caseRow.case_opened_date));

    if (!referralSourceKey || !metricMonth) return;

    const groupKey = [referralSourceKey, metricMonth].join('|');

    if (!out[groupKey]) {
      out[groupKey] = 0;
    }

    out[groupKey] += toNumber_(caseRow.retainer);
  });

  return out;
}

function aggregateConsultationFeesByReferralSourceAndMonth_(caseMasterRows) {
  const out = {};

  caseMasterRows.forEach(function(caseRow) {
    const referralSource = String(firstNonEmpty_(caseRow.lead_referral_source)).trim();
    const referralSourceKey = normalizeText_(referralSource);
    const metricMonth = formatCaseProfitabilityEntryMonth_(firstNonEmpty_(caseRow.case_opened_date));
    const consultationFee = toNumber_(caseRow.consultation_fee);

    if (!referralSourceKey || !metricMonth || !consultationFee) return;

    const groupKey = [referralSourceKey, metricMonth].join('|');

    if (!out[groupKey]) {
      out[groupKey] = {
        consultation_fee_amount: 0,
        consultation_fee_count: 0
      };
    }

    out[groupKey].consultation_fee_amount += consultationFee;
    out[groupKey].consultation_fee_count += 1;
  });

  return out;
}

function buildFactCaseProfitability() {
  const expenses = readSheetAsObjectsIfExists_(CONFIG.sheets.rawExpenses);
  const caseMasterRows = readSheetAsObjectsIfExists_(CONFIG.sheets.factCaseMaster);
  const retainerByReferralSourceAndMonth = aggregateRetainersByReferralSourceAndMonth_(caseMasterRows);
  const consultationFeesByReferralSourceAndMonth = aggregateConsultationFeesByReferralSourceAndMonth_(caseMasterRows);

  if (!expenses.length) {
    writeRowsToSheet_(CONFIG.sheets.factCaseProfitability, []);
    return;
  }

  const grouped = {};

  expenses.forEach(function(expenseRow) {
    const activityName = String(firstNonEmpty_(expenseRow.activity_name, 'Unclassified')).trim() || 'Unclassified';
    if (!isAllowedCaseProfitabilityActivityName_(activityName)) return;

    const entryDate = extractCaseProfitabilityEntryDate_(expenseRow);
    const entryMonth = formatCaseProfitabilityEntryMonth_(entryDate);
    const groupKey = [normalizeText_(activityName), entryDate].join('|');

    if (!grouped[groupKey]) {
      grouped[groupKey] = {
        activity_name: activityName,
        referral_source_linked: getLinkedReferralSourceByActivityName_(activityName),
        entry_date: entryDate,
        entry_month: entryMonth,
        expense_count: 0,
        expense_amount: 0
      };
    }

    grouped[groupKey].expense_count += 1;
    grouped[groupKey].expense_amount += extractCaseProfitabilityExpenseAmount_(expenseRow);
  });

  const rows = Object.keys(grouped)
    .sort()
    .map(function(groupKey) {
      const row = grouped[groupKey];
      const referralSourceKey = normalizeText_(row.referral_source_linked);
      const revenueKey = [referralSourceKey, row.entry_month].join('|');

      row.revenue_associated = referralSourceKey
        ? toNumber_(retainerByReferralSourceAndMonth[revenueKey])
        : 0;
      row.consultation_fee = referralSourceKey && consultationFeesByReferralSourceAndMonth[revenueKey]
        ? toNumber_(consultationFeesByReferralSourceAndMonth[revenueKey].consultation_fee_amount)
        : 0;

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

  ['expense_count', 'expense_amount', 'revenue_associated', 'consultation_fee'].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('0.00');
    }
  });
}
