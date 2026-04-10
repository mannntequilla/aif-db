function getSyncDefinitions_() {
  return {
    cases: {
      endpoint: CONFIG.endpoints.cases,
      sheetName: CONFIG.sheets.rawCases,
      transform: transformCasesForSync_
    },
    clients: {
      endpoint: CONFIG.endpoints.clients,
      sheetName: CONFIG.sheets.rawClients,
      transform: transformClientsForSync_
    },
    events: {
      endpoint: CONFIG.endpoints.events,
      sheetName: CONFIG.sheets.rawEvents,
      transform: passthroughSyncRows_
    },
    invoices: {
      endpoint: CONFIG.endpoints.invoices,
      sheetName: CONFIG.sheets.rawInvoices,
      transform: passthroughSyncRows_
    },
    expenses: {
      endpoint: CONFIG.endpoints.expenses,
      sheetName: CONFIG.sheets.rawExpenses,
      transform: passthroughSyncRows_
    },
    leads: {
      endpoint: CONFIG.endpoints.leads,
      sheetName: CONFIG.sheets.rawLeads,
      transform: passthroughSyncRows_
    },
    roles: {
      endpoint: CONFIG.endpoints.roles,
      sheetName: CONFIG.sheets.rawRoles,
      transform: passthroughSyncRows_
    },
    calls: {
      endpoint: CONFIG.endpoints.calls,
      sheetName: CONFIG.sheets.rawCalls,
      transform: passthroughSyncRows_,
      afterWrite: logSyncedRowsCount_
    },
    tasks: {
      endpoint: CONFIG.endpoints.tasks,
      sheetName: CONFIG.sheets.rawTasks,
      transform: passthroughSyncRows_,
      afterWrite: logSyncedRowsCount_
    },
    staff: {
      endpoint: CONFIG.endpoints.staff,
      sheetName: CONFIG.sheets.rawStaff,
      transform: passthroughSyncRows_
    },
    customFields: {
      endpoint: CONFIG.endpoints.customFields,
      sheetName: CONFIG.sheets.rawCustomFields,
      transform: passthroughSyncRows_
    }
  };
}
