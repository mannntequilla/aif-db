function testCasesFetch() {
  const cases = apiGetAllPages_(CONFIG.endpoints.cases);
  Logger.log('Total cases: ' + cases.length);
}
function resetAutoRefreshTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  // elimina triggers anteriores
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runFullRefreshCaseMaster') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // crea uno nuevo (1 vez al dia)
  ScriptApp.newTrigger('runFullRefreshCaseMaster')
    .timeBased()
    .everyDays(1)
    .create();

  Logger.log('Trigger configurado correctamente.');
}

function runFullRefreshCaseMaster() {
  fullRefreshCaseMaster();
}


function syncAllRaw() {
  syncResourcesByKeys_([
    'cases',
    'clients',
    'leads',
    'invoices',
    'expenses',
    'events',
    'roles',
    'calls',
    'tasks',
    'staff',
    'customFields',
    'referralSources'
  ]);
}


function syncCaseMasterInputs() {
  syncResourcesByKeys_([
    'cases',
    'clients',
    'invoices',
    'expenses',
    'events',
    'customFields'
  ]);
}

/**
 * Refreshes the current case-to-worker assignments independently from the
 * regular case refresh. This keeps the high-volume spreadsheet writes apart
 * and avoids a Sheets timeout after the core reporting tables are rebuilt.
 */
function refreshCaseStaffingAnalytics() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    Logger.log('Ya hay una ejecucion en curso. Intenta nuevamente en unos minutos.');
    return;
  }

  try {
    Logger.log('=== INICIO refreshCaseStaffingAnalytics ===');
    syncResourcesByKeys_(['cases', 'staff']);
    buildBridgeCaseStaff();
    buildOpenCasesByParalegalReport();
    Logger.log('=== FIN OK refreshCaseStaffingAnalytics ===');
  } catch (error) {
    Logger.log('ERROR en refreshCaseStaffingAnalytics: ' + error.message);
    Logger.log(error.stack);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function exploreExpensesRaw() {
  syncExpenses();
  profileExpensesRaw_();
}

function fullRefreshCaseMaster() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    Logger.log('Ya hay una ejecucion en curso.');
    return;
  }

  const start = new Date();

  try {
    Logger.log('=== INICIO fullRefreshCaseMaster ===');

    Logger.log('1. Sync case master inputs...');
    syncCaseMasterInputs();

    Logger.log('2. Import latest MyCase leads report...');
    importLatestMyCaseLeadsReportFromDrive();

    Logger.log('3. Build dim_date...');
    buildDimDate();

    Logger.log('4. Build fact_case_master...');
    buildFactCaseMaster();

    Logger.log('5. Build fact_case...');
    buildFactCase();

    Logger.log('6. Build bridge_client_cases...');
    buildBridgeClientCases();

    Logger.log('7. Build bridge_lead_case...');
    buildBridgeLeadCase();

    Logger.log('8. Build EventsPerCaseId...');
    buildEventsPerCaseId();

    Logger.log('9. Build fact_case_profitability...');
    buildFactCaseProfitability();

    Logger.log('10. updateLastRefreshTimestamp_');
    updateLastRefreshTimestamp_();

    Logger.log('=== FIN OK fullRefreshCaseMaster ===');
    Logger.log('Duracion total: ' + ((new Date() - start) / 1000) + ' segundos');
  } catch (error) {
    Logger.log('ERROR en fullRefreshCaseMaster: ' + error.message);
    Logger.log(error.stack);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function fullRefreshAll() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    Logger.log('Ya hay una ejecucion en curso.');
    return;
  }

  const start = new Date();

  try {
    Logger.log('=== INICIO fullRefreshAll ===');

    Logger.log('1. Sync all raw sheets...');
    syncAllRaw();

    Logger.log('2. Import latest MyCase leads report...');
    importLatestMyCaseLeadsReportFromDrive();

    Logger.log('3. Build dim_date...');
    buildDimDate();

    Logger.log('4. Build fact_case_master...');
    buildFactCaseMaster();

    Logger.log('5. Build fact_case...');
    buildFactCase();

    Logger.log('6. Build bridge_client_cases...');
    buildBridgeClientCases();

    Logger.log('7. Build bridge_lead_case...');
    buildBridgeLeadCase();

    Logger.log('8. Build EventsPerCaseId...');
    buildEventsPerCaseId();

    Logger.log('9. Build fact_case_profitability...');
    buildFactCaseProfitability();

    Logger.log('10. Build case staff table...');
    buildCaseStaffTable();

    Logger.log('11. updateLastRefreshTimestamp_');
    updateLastRefreshTimestamp_();

    Logger.log('=== FIN OK fullRefreshAll ===');
    Logger.log('Duracion total: ' + ((new Date() - start) / 1000) + ' segundos');
  } catch (error) {
    Logger.log('ERROR en fullRefreshAll: ' + error.message);
    Logger.log(error.stack);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function refreshMyCaseLeadsReport(){
  importLatestMyCaseLeadsReportFromDrive()
}

function updateLastRefreshTimestamp_() {
  const sheet = getSpreadsheet_().getSheetByName('Menu');

  if (!sheet) return;

  sheet.getRange('A1').setValue('Ultima actualizacion: ' + new Date());
}
