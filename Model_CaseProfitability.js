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

function getCaseProfitabilityActivitySourceMap_() {
  return {
    'Belgrano': 'Belgrano',
    'Spanish Smile': 'Spanish Smile',
    'El abogado': 'El Abogado.com',
    'Facebook Ads': 'Facebook Ads',
    'Google Ads': 'Google Ads',
    'Instagram': 'Instagram',
    'Website': 'Website'
  };
}

function normalizeCaseProfitabilityActivityName_(activityName) {
  const rawName = String(firstNonEmpty_(activityName, 'Unclassified')).trim() || 'Unclassified';
  const normalizedActivityName = normalizeText_(rawName);

  if (normalizedActivityName.indexOf(normalizeText_('Belgrano')) !== -1) return 'Belgrano';
  if (normalizedActivityName.indexOf(normalizeText_('Spanish Smile')) !== -1) return 'Spanish Smile';
  if (normalizedActivityName.indexOf(normalizeText_('Facebook Ads')) !== -1) return 'Facebook Ads';
  if (normalizedActivityName.indexOf(normalizeText_('Google Ads')) !== -1) return 'Google Ads';
  if (normalizedActivityName.indexOf(normalizeText_('Instagram')) !== -1) return 'Instagram';
  if (normalizedActivityName.indexOf(normalizeText_('Website')) !== -1) return 'Website';

  if (
    normalizedActivityName.indexOf(normalizeText_('El abogado')) !== -1 ||
    normalizedActivityName.indexOf(normalizeText_('Abogado')) !== -1
  ) {
    return 'El abogado';
  }

  return rawName;
}

function getLinkedReferralSourceByActivityName_(activityName) {
  const sourceMap = getCaseProfitabilityActivitySourceMap_();
  return firstNonEmpty_(sourceMap[activityName], '');
}

function aggregateMetricByReferralSourceAndMonth_(caseMasterRows, valueGetter) {
  const out = {};

  caseMasterRows.forEach(function(caseRow) {
    const referralSource = String(firstNonEmpty_(caseRow.lead_referral_source)).trim();
    const referralSourceKey = normalizeText_(referralSource);
    const metricMonth = formatCaseProfitabilityEntryMonth_(firstNonEmpty_(caseRow.case_opened_date));

    if (!referralSourceKey || !metricMonth) return;

    const groupKey = [referralSourceKey, metricMonth].join('|');
    out[groupKey] = (out[groupKey] || 0) + toNumber_(valueGetter(caseRow));
  });

  return out;
}

function aggregateConsultationMetricsByReferralSourceAndMonth_(caseMasterRows) {
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
        consultation_fee: 0,
        consultation_count: 0
      };
    }

    out[groupKey].consultation_fee += consultationFee;
    out[groupKey].consultation_count += 1;
  });

  return out;
}

function getMetricValueByReferralSourceAndMonth_(metricIndex, referralSource, entryMonth) {
  if (!referralSource || !entryMonth) return 0;

  const key = [normalizeText_(referralSource), entryMonth].join('|');
  return toNumber_(metricIndex[key]);
}

function getConsultationMetricByReferralSourceAndMonth_(consultationIndex, referralSource, entryMonth, fieldName) {
  if (!referralSource || !entryMonth) return 0;

  const key = [normalizeText_(referralSource), entryMonth].join('|');
  const metricRow = consultationIndex[key];
  return toNumber_(metricRow && metricRow[fieldName]);
}

function buildFactCaseProfitability() {
  const expenses = readSheetAsObjectsIfExists_(CONFIG.sheets.rawExpenses);
  const caseMasterRows = readSheetAsObjectsIfExists_(CONFIG.sheets.factCaseMaster);
  const retainerByReferralSourceAndMonth = aggregateMetricByReferralSourceAndMonth_(caseMasterRows, function(caseRow) {
    return caseRow.retainer;
  });
  const totalRevenueByReferralSourceAndMonth = aggregateMetricByReferralSourceAndMonth_(caseMasterRows, function(caseRow) {
    return caseRow.total_invoice_amount;
  });
  const consultationByReferralSourceAndMonth = aggregateConsultationMetricsByReferralSourceAndMonth_(caseMasterRows);

  if (!expenses.length) {
    writeRowsToSheet_(CONFIG.sheets.factCaseProfitability, []);
    return;
  }

  const grouped = {};

  expenses.forEach(function(expenseRow) {
    const activityName = normalizeCaseProfitabilityActivityName_(expenseRow.activity_name);
    const referralSourceLinked = getLinkedReferralSourceByActivityName_(activityName);
    if (!referralSourceLinked) return;

    const entryDate = extractCaseProfitabilityEntryDate_(expenseRow);
    const entryMonth = formatCaseProfitabilityEntryMonth_(entryDate);
    const groupKey = [normalizeText_(activityName), entryDate].join('|');

    if (!grouped[groupKey]) {
      grouped[groupKey] = {
        activity_name: activityName,
        referral_source_linked: referralSourceLinked,
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

      row.retainer_revenue = getMetricValueByReferralSourceAndMonth_(
        retainerByReferralSourceAndMonth,
        row.referral_source_linked,
        row.entry_month
      );
      row.total_revenue = getMetricValueByReferralSourceAndMonth_(
        totalRevenueByReferralSourceAndMonth,
        row.referral_source_linked,
        row.entry_month
      );
      row.consultation_count = getConsultationMetricByReferralSourceAndMonth_(
        consultationByReferralSourceAndMonth,
        row.referral_source_linked,
        row.entry_month,
        'consultation_count'
      );
      row.consultation_fee = getConsultationMetricByReferralSourceAndMonth_(
        consultationByReferralSourceAndMonth,
        row.referral_source_linked,
        row.entry_month,
        'consultation_fee'
      );

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

  ['expense_count', 'expense_amount', 'retainer_revenue', 'total_revenue', 'consultation_count', 'consultation_fee'].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('0.00');
    }
  });
}
