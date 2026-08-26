/**
 * Controlled current-stage classification for the Operations Overview.
 * The values reflect the agreed operational meaning of each MyCase stage;
 * they do not imply historical stage movement or individual performance.
 */
function buildStageOperatingModel() {
  const rows = [
    stageOperatingRow_('Recently Added', 'Active Work'),
    stageOperatingRow_('Case Classification', 'Active Work'),
    stageOperatingRow_('Drafting Packet', 'Active Work'),
    stageOperatingRow_('Case Assembly & Filing', 'Active Work'),
    stageOperatingRow_('Court Preparation', 'Active Work'),

    stageOperatingRow_('Compliance Review', 'Active Follow Up',
      'Verification that forms and submitted documents are complete, correct and appropriate'),

    stageOperatingRow_('Welcome Letter Sent', 'Client Waiting'),
    stageOperatingRow_('Document Gathering', 'Active Work',
      'Active collection and coordination of client documents'),
    stageOperatingRow_('Client Review & Signature', 'Active Work',
      'Active work; client corrections and interactions are generally minimal'),

    stageOperatingRow_('Waiting for Official Response', 'External Waiting',
      'May require client contact or an inquiry during processing time',
      { requiresPeriodicFollowUp: true }),
    stageOperatingRow_('Waiting for Final Hearing', 'External Waiting',
      'May require client contact or an inquiry while awaiting hearing',
      { requiresPeriodicFollowUp: true }),
    stageOperatingRow_('Waiting Priority Date', 'External Waiting',
      'May require periodic follow-up while awaiting priority date',
      { requiresPeriodicFollowUp: true }),

    stageOperatingRow_('Final Resolution', 'Closure / Exit'),
    stageOperatingRow_('Disengagement', 'Closure / Exit',
      'Close after the 30-day disengagement period',
      { requiresCaseClosure: false, closureGraceDays: 30 }),

    stageOperatingRow_('Active', 'Unclassified / Data Quality', 'Stage update required')
  ];

  writeRowsToSheet_(CONFIG.sheets.stageOperatingModel, rows);
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.sheets.stageOperatingModel);
  if (sheet) sheet.setFrozenRows(1);
}

function stageOperatingRow_(stageName, operatingCategory, note, options) {
  options = options || {};
  const isActiveWork = operatingCategory === 'Active Work';
  const isActiveFollowUp = operatingCategory === 'Active Follow Up';
  const isWaiting = operatingCategory === 'Client Waiting' ||
    operatingCategory === 'External Waiting' ||
    operatingCategory === 'Attorney Review';
  const isClosure = operatingCategory === 'Closure / Exit';
  const isDataQuality = operatingCategory === 'Unclassified / Data Quality';

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
