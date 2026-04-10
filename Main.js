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

  // crea uno nuevo (cada 1 hora)
  ScriptApp.newTrigger('runFullRefreshCaseMaster')
    .timeBased()
    .everyMinutes(15)
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
    'customFields'
  ]);
}


function syncCaseMasterInputs() {
  syncResourcesByKeys_([
    'cases',
    'clients',
    'invoices',
    'events',
    'customFields'
  ]);
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

    Logger.log('3. Build fact_case_master...');
    buildFactCaseMaster();

    Logger.log('4. updateLastRefreshTimestamp_');
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

function refreshMyCaseLeadsReport(){
  importLatestMyCaseLeadsReportFromDrive()
}

function updateLastRefreshTimestamp_() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Menu');

  if (!sheet) return;

  sheet.getRange('A1').setValue('Ultima actualizacion: ' + new Date());
}
