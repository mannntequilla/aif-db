function syncResourceByKey_(resourceKey) {
  const definitions = getSyncDefinitions_();
  const definition = definitions[resourceKey];

  if (!definition) {
    throw new Error('Unknown sync resource: ' + resourceKey);
  }

  const rows = apiGetAllPages_(definition.endpoint);
  const transform = definition.transform || passthroughSyncRows_;
  const normalizedRows = transform(rows);

  writeRowsToSheet_(definition.sheetName, normalizedRows);

  if (typeof definition.afterWrite === 'function') {
    definition.afterWrite(normalizedRows, definition);
  }

  return normalizedRows;
}

function syncResourcesByKeys_(resourceKeys) {
  return resourceKeys.map(function(resourceKey) {
    return {
      key: resourceKey,
      rows: syncResourceByKey_(resourceKey)
    };
  });
}

function logSyncedRowsCount_(rows) {
  Logger.log('Rows synced: ' + rows.length);
}
