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

    stageOperatingRow_('Welcome Letter Sent', 'Client Waiting'),
    stageOperatingRow_('Document Gathering', 'Client Waiting'),
    stageOperatingRow_('Client Review & Signature', 'Client Waiting'),

    stageOperatingRow_('Waiting for Official Response', 'External Waiting'),
    stageOperatingRow_('Waiting for Final Hearing', 'External Waiting'),
    stageOperatingRow_('Waiting Priority Date', 'External Waiting'),

    stageOperatingRow_('Final Resolution', 'Closure / Exit'),
    stageOperatingRow_('Disengagement', 'Closure / Exit'),

    stageOperatingRow_('Active', 'Unclassified / Data Quality', 'Stage update required')
  ];

  writeRowsToSheet_(CONFIG.sheets.stageOperatingModel, rows);
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.sheets.stageOperatingModel);
  if (sheet) sheet.setFrozenRows(1);
}

function stageOperatingRow_(stageName, operatingCategory, note) {
  const isActiveWork = operatingCategory === 'Active Work';
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
    is_waiting: isWaiting ? 1 : 0,
    is_closure_exit: isClosure ? 1 : 0,
    is_valid_for_active_work_share: (isActiveWork || isWaiting) ? 1 : 0,
    requires_stage_correction: isDataQuality ? 1 : 0,
    requires_case_closure: isClosure ? 1 : 0,
    operating_note: note || ''
  };
}
