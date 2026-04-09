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
  syncCases();
  syncClients();
  syncLeads();
  syncInvoices();
  syncEvents();
  syncRoles();
  syncCalls();
  syncTasks();
}


function syncCaseMasterInputs() {
  syncCases();
  syncClients();
  syncInvoices();
  syncEvents();
}

function fullRefreshCaseMaster() {
  syncCaseMasterInputs();
  importLatestMyCaseLeadsReportFromDrive();
  buildFactCaseMaster();
}

function refreshMyCaseLeadsReport(){
  importLatestMyCaseLeadsReportFromDrive()
}