function transformCasesForSync_(rows) {
  return rows.map(function(caseRow) {
    const office = caseRow.office || {};

    return Object.assign({}, caseRow, {
      office: JSON.stringify(caseRow.office || {}),
      office_name: office.name || '',
      office_id: office.id || '',
      clients: stringifyIdsDeep_(caseRow.clients || []),
      custom_field_values: stringifyIdsDeep_(caseRow.custom_field_values || [])
    });
  });
}

function transformClientsForSync_(rows) {
  return rows.map(function(client) {
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
}

function passthroughSyncRows_(rows) {
  return rows;
}
