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

function getLinkedReferralSourcesByActivityName_(activityName) {
  const normalizedActivityName = normalizeText_(activityName);

  if (normalizedActivityName === normalizeText_('Belgrano')) {
    return ['Google Ads', 'Website'];
  }

  if (normalizedActivityName === normalizeText_('Spanish Smile')) {
    return ['Spanish Smile'];
  }

  if (normalizedActivityName === normalizeText_('El abogado')) {
    return ['El Abogado.com'];
  }

  if (normalizedActivityName === normalizeText_('Facebook Ads')) {
    return ['Spanish Smile'];
  }

  return [];
}

function formatLinkedReferralSources_(referralSources) {
  return referralSources.join(', ');
}

function sumRetainersForReferralSourcesAndMonth_(retainersByReferralSourceAndMonth, referralSources, entryMonth) {
  return referralSources.reduce(function(total, referralSource) {
    const revenueKey = [normalizeText_(referralSource), entryMonth].join('|');
    return total + toNumber_(retainersByReferralSourceAndMonth[revenueKey]);
  }, 0);
}

function aggregateTotalRevenueByReferralSourceAndMonth_(caseMasterRows) {
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

    out[groupKey] += toNumber_(caseRow.total_invoice_amount);
  });

  return out;
}

function sumTotalRevenueForReferralSourcesAndMonth_(totalRevenueByReferralSourceAndMonth, referralSources, entryMonth) {
  return referralSources.reduce(function(total, referralSource) {
    const revenueKey = [normalizeText_(referralSource), entryMonth].join('|');
    return total + toNumber_(totalRevenueByReferralSourceAndMonth[revenueKey]);
  }, 0);
}

function sumConsultationFeesForReferralSourcesAndMonth_(consultationFeesByReferralSourceAndMonth, referralSources, entryMonth) {
  return referralSources.reduce(function(total, referralSource) {
    const consultationKey = [normalizeText_(referralSource), entryMonth].join('|');
    const consultationData = consultationFeesByReferralSourceAndMonth[consultationKey];
    return total + toNumber_(consultationData && consultationData.consultation_fee_amount);
  }, 0);
}

function sumConsultationCountsForReferralSourcesAndMonth_(consultationFeesByReferralSourceAndMonth, referralSources, entryMonth) {
  return referralSources.reduce(function(total, referralSource) {
    const consultationKey = [normalizeText_(referralSource), entryMonth].join('|');
    const consultationData = consultationFeesByReferralSourceAndMonth[consultationKey];
    return total + toNumber_(consultationData && consultationData.consultation_fee_count);
  }, 0);
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
  const totalRevenueByReferralSourceAndMonth = aggregateTotalRevenueByReferralSourceAndMonth_(caseMasterRows);
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
      const linkedReferralSources = getLinkedReferralSourcesByActivityName_(activityName);

      grouped[groupKey] = {
        activity_name: activityName,
        referral_source_linked: formatLinkedReferralSources_(linkedReferralSources),
        entry_date: entryDate,
        entry_month: entryMonth,
        expense_count: 0,
        expense_amount: 0,
        linked_referral_sources_raw: linkedReferralSources
      };
    }

    grouped[groupKey].expense_count += 1;
    grouped[groupKey].expense_amount += extractCaseProfitabilityExpenseAmount_(expenseRow);
  });

  const rows = Object.keys(grouped)
    .sort()
    .map(function(groupKey) {
      const row = grouped[groupKey];
      const linkedReferralSources = row.linked_referral_sources_raw || [];

      row.retainer_revenue = linkedReferralSources.length
        ? sumRetainersForReferralSourcesAndMonth_(
            retainerByReferralSourceAndMonth,
            linkedReferralSources,
            row.entry_month
          )
        : 0;
      row.total_revenue = linkedReferralSources.length
        ? sumTotalRevenueForReferralSourcesAndMonth_(
            totalRevenueByReferralSourceAndMonth,
            linkedReferralSources,
            row.entry_month
          )
        : 0;
      row.consultation_count = linkedReferralSources.length
        ? sumConsultationCountsForReferralSourcesAndMonth_(
            consultationFeesByReferralSourceAndMonth,
            linkedReferralSources,
            row.entry_month
          )
        : 0;
      row.consultation_fee = linkedReferralSources.length
        ? sumConsultationFeesForReferralSourcesAndMonth_(
            consultationFeesByReferralSourceAndMonth,
            linkedReferralSources,
            row.entry_month
          )
        : 0;

      delete row.linked_referral_sources_raw;

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
