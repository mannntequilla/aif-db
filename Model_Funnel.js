function buildLeadsFunnelByDate() {
  const rows = readSheetAsObjectsIfExists_(CONFIG.sheets.rawMyCaseLeadsReport);

  if (!rows || !rows.length) {
    writeRowsToSheet_(CONFIG.sheets.funnelLeadsByDate, []);
    return;
  }

  const grouped = {};

  rows.forEach(function(row) {
    const funnelDate = toDateOnlyMaybe_(firstNonEmpty_(row['Date added']));
    if (!funnelDate) return;

    const rawStatus = firstNonEmpty_(row['Lead status']);
    const status = rawStatus ? String(rawStatus).trim() : '';

    const conversionDate = toDateOnlyMaybe_(firstNonEmpty_(row['Conversion date']));
    const stage = classifyLeadFunnelStage_(status, conversionDate);

    if (!grouped[funnelDate]) {
      grouped[funnelDate] = {
        'New Leads': 0,
        'Potential Leads': 0,
        'Converted': 0
      };
    }

    grouped[funnelDate]['New Leads'] += 1;

    if (stage === 'Potential Leads') {
      grouped[funnelDate]['Potential Leads'] += 1;
    }

    if (stage === 'Converted') {
      grouped[funnelDate]['Potential Leads'] += 1;
      grouped[funnelDate]['Converted'] += 1;
    }
  });

  const stageOrderMap = {
    'New Leads': 1,
    'Potential Leads': 2,
    'Converted': 3
  };

  const output = [];

  Object.keys(grouped)
    .sort()
    .forEach(function(date) {
      Object.keys(stageOrderMap).forEach(function(stage) {
        output.push({
          funnel_date: formatDateOnlyForSheet_(date),
          stage: stage,
          count: grouped[date][stage],
          stage_order: stageOrderMap[stage]
        });
      });
    });

  writeRowsToSheet_(CONFIG.sheets.funnelLeadsByDate, output);
}

function classifyLeadFunnelStage_(status, conversionDate) {
  const normalizedStatus = normalizeLeadStatus_(status);

  const isConvertedStatus =
    normalizedStatus === 'contract' ||
    normalizedStatus === 'detainee visitation';

  if (isConvertedStatus && conversionDate) {
    return 'Converted';
  }

  const isPotentialLead =
    normalizedStatus === 'consult scheduled' ||
    normalizedStatus === 'hot deal' ||
    normalizedStatus === 'first follow up' ||
    normalizedStatus === 'second follow up';

  if (isPotentialLead) {
    return 'Potential Leads';
  }

  const isNewLead =
    normalizedStatus === 'new lead' ||
    normalizedStatus === 'pending payment' ||
    normalizedStatus === 'contacted' ||
    normalizedStatus === 'contacted 2nd follow up' ||
    normalizedStatus === 'contacted 3rd follow up';

  if (isNewLead) {
    return 'New Leads';
  }

  return 'New Leads';
}

function normalizeLeadStatus_(status) {
  return String(status || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}
