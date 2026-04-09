function syncCases() {
  const rows = apiGetAllPages_(CONFIG.endpoints.cases);

  const normalized = rows.map(function(caseRow) {
    const office = caseRow.office || {};
    const clients = stringifyIdsDeep_(caseRow.clients || []);
    const customFieldValues = stringifyIdsDeep_(caseRow.custom_field_values || []);

    return Object.assign({}, caseRow, {
      office: JSON.stringify(caseRow.office || {}),
      office_name: office.name || '',
      office_id: office.id || ''
    });
  });

  writeRowsToSheet_(CONFIG.sheets.rawCases, normalized);
}


function syncClients() {
  const rows = apiGetAllPages_(CONFIG.endpoints.clients);

  const normalized = rows.map(function(client) {
    const address = client.address || {};

    return Object.assign({}, client, {
      address: JSON.stringify(client.address || {}),
      address1: address.address1 || '',
      address2: address.address2 || '',
      city: address.city || '',
      state: address.state || '',
      zip_code: address.zip_code || '',
      country: address.country || ''
    });
  });

  writeRowsToSheet_(CONFIG.sheets.rawClients, normalized);
}

function syncEvents() {
  const rows = apiGetAllPages_(CONFIG.endpoints.events);
  writeRowsToSheet_(CONFIG.sheets.rawEvents, rows);
}

function syncInvoices() {
  const rows = apiGetAllPages_(CONFIG.endpoints.invoices);
  writeRowsToSheet_(CONFIG.sheets.rawInvoices, rows);
}

function syncLeads() {
  const rows = apiGetAllPages_(CONFIG.endpoints.leads);
  writeRowsToSheet_(CONFIG.sheets.rawLeads, rows);
}

function syncRoles() {
  const rows = apiGetAllPages_(CONFIG.endpoints.roles);
  writeRowsToSheet_(CONFIG.sheets.rawRoles, rows);
}

function syncCalls() {
  const rows = apiGetAllPages_(CONFIG.endpoints.calls);
  writeRowsToSheet_(CONFIG.sheets.rawCalls, rows);
  console.log(rows.length);
}

function syncTasks(){
  const rows = apiGetAllPages_(CONFIG.endpoints.tasks);
  writeRowsToSheet_(CONFIG.sheets.rawTasks, rows);
  console.log(rows.length);
}

function syncStaff (){
  const rows = apiGetAllPages_(CONFIG.endpoints.staff);
  writeRowsToSheet_(CONFIG.sheets.rawStaff, rows);
}