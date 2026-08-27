/**
 * Controlled current-stage classification for the Operations Overview.
 * The values reflect the agreed operational meaning of each MyCase stage;
 * they do not imply historical stage movement or individual performance.
 */
function buildStageOperatingModel() {
  const rows = [
    stageOperatingRow_('Recently Added', 'Filing Pending', '', { classification: 'Active Work' }),
    stageOperatingRow_('Case Classification', 'Filing Pending', '', { classification: 'Active Work' }),
    stageOperatingRow_('Drafting Packet', 'Filing Pending', '', { classification: 'Active Work' }),
    stageOperatingRow_('Case Assembly & Filing', 'Filing Pending', '', { classification: 'Active Work' }),
    stageOperatingRow_('Court Preparation', 'Filing Pending', '', { classification: 'Active Work' }),

    stageOperatingRow_('Welcome Letter Sent', 'Pending Signature', '', { classification: 'Client Waiting' }),
    stageOperatingRow_('Document Gathering', 'Filing Pending',
      'Active collection and coordination of client documents', { classification: 'Active Follow Up' }),
    stageOperatingRow_('Client Review & Signature', 'Filing Pending',
      'Active work; client corrections and interactions are generally minimal', { classification: 'Active Follow Up' }),

    stageOperatingRow_('Waiting for Official Response', 'Filed',
      'May require client contact or an inquiry during processing time',
      { classification: 'External Waiting', requiresPeriodicFollowUp: true }),
    stageOperatingRow_('Waiting for Final Hearing', 'Filed',
      'May require client contact or an inquiry while awaiting hearing',
      { classification: 'External Waiting', requiresPeriodicFollowUp: true }),
    stageOperatingRow_('Waiting Priority Date', 'Filed',
      'May require periodic follow-up while awaiting priority date',
      { classification: 'External Waiting', requiresPeriodicFollowUp: true }),

    stageOperatingRow_('Final Resolution', 'File Closed', '', { classification: 'Closure / Exit' }),
    stageOperatingRow_('Disengagement', 'File Closed',
      'Close after the 30-day disengagement period',
      { classification: 'Closure / Exit', requiresCaseClosure: false, closureGraceDays: 30 }),

    stageOperatingRow_('Active', 'Unclassified / Data Quality', 'Stage update required')
  ];

  writeRowsToSheet_(CONFIG.sheets.stageOperatingModel, rows);
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.sheets.stageOperatingModel);
  if (sheet) sheet.setFrozenRows(1);
}

function stageOperatingRow_(stageName, operatingCategory, note, options) {
  options = options || {};
  const classification = options.classification || operatingCategory;
  const isActiveWork = classification === 'Active Work';
  const isActiveFollowUp = classification === 'Active Follow Up';
  const isWaiting = classification === 'Client Waiting' ||
    classification === 'External Waiting' ||
    classification === 'Attorney Review';
  const isClosure = classification === 'Closure / Exit';
  const isDataQuality = classification === 'Unclassified / Data Quality';

  return {
    current_case_stage_key: normalizeDimensionKey_(stageName),
    current_case_stage: stageName,
    operating_category: operatingCategory,
    is_active_work: isActiveWork ? 1 : 0,
    is_active_follow_up: isActiveFollowUp ? 1 : 0,
    is_waiting: isWaiting ? 1 : 0,
    is_closure_exit: isClosure ? 1 : 0,
    is_valid_for_active_work_share: (isActiveWork || isActiveFollowUp || isWaiting) ? 1 : 0,
    requires_stage_correction: isDataQuality ? 1 : 0,
    requires_case_closure: options.requiresCaseClosure === false ? 0 : (isClosure ? 1 : 0),
    closure_grace_days: options.closureGraceDays || 0,
    requires_periodic_follow_up: options.requiresPeriodicFollowUp ? 1 : 0,
    operating_note: note || ''
  };
}
