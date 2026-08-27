/**
 * Main daily entrypoint. It refreshes every source currently used by the
 * Looker Studio report and rebuilds its dependent reporting tables.
 */
function runDailyRefresh() {
  refreshReporting_();
}

/**
 * One-time administrative action. Run this after deploying the cleanup to
 * replace the former runFullRefreshCaseMaster trigger with runDailyRefresh.
 */
function resetDailyRefreshTrigger() {
  const triggerHandlers = {
    runFullRefreshCaseMaster: true,
    runDailyRefresh: true
  };

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (triggerHandlers[trigger.getHandlerFunction()]) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('runDailyRefresh')
    .timeBased()
    .atHour(15)
    .nearMinute(0)
    .everyDays(1)
    .inTimezone('America/New_York')
    .create();

  Logger.log('Trigger diario configurado para runDailyRefresh: 3:00 p. m. America/New_York.');
}

/**
 * Refreshes the current case-to-worker workload model separately from the
 * daily reporting refresh, preventing accumulated Sheets write timeouts.
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
    buildCaseWorkloadByStaff();
    Logger.log('=== FIN OK refreshCaseStaffingAnalytics ===');
  } catch (error) {
    Logger.log('ERROR en refreshCaseStaffingAnalytics: ' + error.message);
    Logger.log(error.stack);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Rebuilds the current workload table used by the Operations Overview.
 */
function refreshOperationsOverview() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    Logger.log('Ya hay una ejecucion en curso. Intenta nuevamente en unos minutos.');
    return;
  }

  try {
    Logger.log('=== INICIO refreshOperationsOverview ===');
    syncResourcesByKeys_(['cases', 'staff']);
    buildStageOperatingModel();
    buildCaseWorkloadByStaff();
    Logger.log('=== FIN OK refreshOperationsOverview ===');
  } catch (error) {
    Logger.log('ERROR en refreshOperationsOverview: ' + error.message);
    Logger.log(error.stack);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function refreshReporting_() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    Logger.log('Ya hay una ejecucion en curso.');
    return;
  }

  const start = new Date();

  try {
    Logger.log('=== INICIO runDailyRefresh ===');

    Logger.log('1. Sync report inputs...');
    syncReportingInputs_();

    Logger.log('2. Import latest MyCase leads report...');
    importLatestMyCaseLeadsReportFromDrive();

    Logger.log('3. Build fact_case_master...');
    buildFactCaseMaster();

    Logger.log('4. Build fact_case...');
    buildFactCase();

    Logger.log('5. Build bridge_client_cases...');
    buildBridgeClientCases();

    Logger.log('6. Build bridge_lead_case...');
    buildBridgeLeadCase();

    Logger.log('7. Build EventsPerCaseId...');
    buildEventsPerCaseId();

    Logger.log('8. Build fact_case_profitability...');
    buildFactCaseProfitability();

    updateLastRefreshTimestamp_();
    Logger.log('=== FIN OK runDailyRefresh ===');
    Logger.log('Duracion total: ' + ((new Date() - start) / 1000) + ' segundos');
  } catch (error) {
    Logger.log('ERROR en runDailyRefresh: ' + error.message);
    Logger.log(error.stack);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function syncReportingInputs_() {
  syncResourcesByKeys_([
    'cases',
    'clients',
    'invoices',
    'expenses',
    'events',
    'customFields'
  ]);
}

function updateLastRefreshTimestamp_() {
  const sheet = getSpreadsheet_().getSheetByName('Menu');
  if (!sheet) return;

  sheet.getRange('A1').setValue('Ultima actualizacion: ' + new Date());
}
