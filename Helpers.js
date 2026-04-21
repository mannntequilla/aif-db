/**
 * Legacy helpers kept temporarily to avoid losing historical code during refactor.
 * Active helpers were moved into focused modules:
 * - CoreSheets.js
 * - CoreObjects.js
 * - CoreDates.js
 * - CoreNormalize.js
 * - CaseMasterHelpers.js
 */

function legacyNormalizePhoneUnused_(value) {
  const digits = String(value || '').replace(/\D+/g, '');

  if (!digits) return '';

  if (digits.length > 10) {
    return digits.slice(-10);
  }

  return digits;
}

function legacyMyCaseGetFullResponseUnused_(path, queryParams) {
  const service = getMyCaseService_();

  if (!service.hasAccess()) {
    throw new Error('Primero debes autorizar la conexion ejecutando beginAuth().');
  }

  let url = `${MYCASE_API_BASE}${path}`;

  if (queryParams) {
    const qs = Object.keys(queryParams)
      .filter(function(key) {
        return queryParams[key] !== null && queryParams[key] !== undefined && queryParams[key] !== '';
      })
      .map(function(key) {
        return `${encodeURIComponent(key)}=${encodeURIComponent(queryParams[key])}`;
      })
      .join('&');

    if (qs) url += `?${qs}`;
  }

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      Authorization: `Bearer ${service.getAccessToken()}`,
      Accept: 'application/json'
    },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error(`Error MyCase ${code}: ${text}`);
  }

  return response;
}

function legacyMyCaseGetUnused_(path, queryParams) {
  const response = legacyMyCaseGetFullResponseUnused_(path, queryParams);
  return JSON.parse(response.getContentText());
}

function legacyExtractNextPageTokenFromLinkUnused_(linkHeader) {
  if (!linkHeader) return null;
  const match = linkHeader.match(/[?&]page_token=([^&>]+)/);
  return match ? decodeURIComponent(match[1]) : null;
}

function legacyGetAllPaginatedUnused_(path, baseParams) {
  let allRecords = [];
  let pageToken = null;
  let keepGoing = true;

  while (keepGoing) {
    const params = Object.assign({}, baseParams || {});
    if (pageToken) {
      params.page_token = pageToken;
    }

    const response = legacyMyCaseGetFullResponseUnused_(path, params);
    const parsed = JSON.parse(response.getContentText());
    const records = Array.isArray(parsed) ? parsed : (parsed.data || []);

    allRecords = allRecords.concat(records);

    const headers = response.getHeaders();
    const linkHeader = headers.Link || headers.link || '';
    pageToken = legacyExtractNextPageTokenFromLinkUnused_(linkHeader);
    keepGoing = !!pageToken;
  }

  return allRecords;
}

function legacyFullRefreshCaseMasterUnused() {
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
    legacyUpdateLastRefreshTimestampUnused_();

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

function legacyUpdateLastRefreshTimestampUnused_() {
  const sheet = getSpreadsheet_().getSheetByName('Menu');

  if (!sheet) return;

  sheet.getRange('A1').setValue('Ultima actualizacion: ' + new Date());
}
