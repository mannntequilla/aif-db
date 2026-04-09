function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet_(name) {
  const ss = getSpreadsheet_();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function clearSheet_(sheetName) {
  const sheet = getOrCreateSheet_(sheetName);
  sheet.clearContents();
  return sheet;
}

function readSheetAsObjectsIfExists_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  const headers = values[0];

  return values.slice(1).map(function(row) {
    const obj = {};
    headers.forEach(function(header, i) {
      obj[header] = row[i];
    });
    return obj;
  });
}

function ensureSheetExists_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

function safeCellValue_(value) {
  if (value === null || value === undefined) return '';
  
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return value;
  }
  
  if (typeof value === 'object') return JSON.stringify(value);
  return value;
}

function writeRowsToSheet_(sheetName, rows) {
  if (!sheetName) {
    throw new Error('Sheet name is missing');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    return sheet;
  }
  sheet = ss.insertSheet();
  sheet.setName(sheetName);

  return sheet;
}

function clearSheet_(sheetName) {
  const sheet = getOrCreateSheet_(sheetName);
  sheet.clearContents();
  return sheet;
}

function writeRowsToSheet_(sheetName, rows) {
  const sheet = clearSheet_(sheetName);

  if (!rows || !rows.length) return;

  const headers = [...new Set(rows.flatMap(row => Object.keys(row)))];

  const values = rows.map(row =>
    headers.map(header => safeCellValue_(row[header]))
  );

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

/************************************************************
HELPERS FOR MODELING
************************************************************/

function readSheetAsObjects_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0];

  return values.slice(1).map(function (row) {
    const obj = {};
    headers.forEach(function (header, i) {
      obj[header] = row[i];
    });
    return obj;
  });
}

function indexBy_(rows, key) {
  const out = {};
  rows.forEach(function (row) {
    const value = firstNonEmpty_(row[key]);
    if (value !== '' && value !== null && value !== undefined) {
      out[String(value)] = row;
    }
  });
  return out;
}

function aggregateInvoicesByCaseId_(invoices) {
  const out = {};

  invoices.forEach(function (inv) {
    const caseId = extractCaseIdFromInvoice_(inv);
    if (!caseId) return;

    const key = String(caseId);
    if (!out[key]) {
      out[key] = {
        total_invoice_amount: 0,
        total_paid_so_far: 0,
        total_balance: 0
      };
    }

    const totalAmount = toNumber_(firstNonEmpty_(inv.total_amount, inv.invoice_total_amount));
    const paidAmount = toNumber_(firstNonEmpty_(inv.paid_amount, inv.total_paid));
    const balance = totalAmount - paidAmount;

    out[key].total_invoice_amount += totalAmount;
    out[key].total_paid_so_far += paidAmount;
    out[key].total_balance += balance;
  });

  return out;
}


function findPreferredCaseClientRef_(caseRow) {
  const candidates = parseJsonMaybe_(firstNonEmpty_(caseRow.clients, caseRow.case_clients, '[]'));

  if (Array.isArray(candidates)) {
    const preferred = candidates.find(function (c) {
      const roleText = JSON.stringify(c).toLowerCase();
      return roleText.indexOf('beneficiary') !== -1 || roleText.indexOf('alien') !== -1;
    });

    if (preferred) return preferred;
    if (candidates.length) return candidates[0];
  }

  return null;
}

function resolveClientFromRef_(clientRef, clientsById) {
  if (!clientRef) return null;

  const clientId = firstNonEmpty_(clientRef.id, clientRef.client_id, clientRef.person_id);
  if (!clientId) return null;

  return clientsById[String(clientId)] || null;
}

function extractCaseIdFromInvoice_(inv) {
  const caseObj = parseJsonMaybe_(inv.case);
  return firstNonEmpty_(
    inv.case_id,
    caseObj && caseObj.id
  );
}

function extractCaseIdFromEvent_(ev) {
  const caseObj = parseJsonMaybe_(ev.case);
  return firstNonEmpty_(
    ev.case_id,
    caseObj && caseObj.id
  );
}

function extractOfficeName_(caseRow) {
  const officeObj = parseJsonMaybe_(caseRow.office);
  return firstNonEmpty_(
    caseRow.office_name,
    officeObj && officeObj.name
  );
}

function extractAddressLine_(client) {
  const address = parseJsonMaybe_(client.address);
  return firstNonEmpty_(
    client.address1,
    address && address.address1,
    client.address
  );
}

function extractCity_(client) {
  const address = parseJsonMaybe_(client.address);
  return firstNonEmpty_(client.city, address && address.city);
}

function extractState_(client) {
  const address = parseJsonMaybe_(client.address);
  return firstNonEmpty_(client.state, address && address.state);
}

function buildFullName_(person) {
  return [
    firstNonEmpty_(person.first_name),
    firstNonEmpty_(person.middle_name),
    firstNonEmpty_(person.last_name)
  ].filter(Boolean).join(' ').trim();
}

function parseJsonMaybe_(value) {
  if (!value) return null;
  if (typeof value === 'object') return value;

  const text = String(value).trim();
  if (!text) return null;

  if ((text.startsWith('{') && text.endsWith('}')) ||
      (text.startsWith('[') && text.endsWith(']'))) {
    try {
      return JSON.parse(text);
    } catch (e) {
      return null;
    }
  }

  return null;
}

function firstNonEmpty_() {
  for (var i = 0; i < arguments.length; i++) {
    var v = arguments[i];
    if (v !== null && v !== undefined && v !== '') return v;
  }
  return '';
}

function toNumber_(value) {
  const n = Number(value || 0);
  return isNaN(n) ? 0 : n;
}

function formatDateOnlyForSheet_(value) {
  if (!value) return '';

  const d = new Date(value);
  if (isNaN(d.getTime())) return '';

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}

function getFirstInitialConsultationByCaseId_(events) {
  const out = {};

  events.forEach(function(ev) {
    const caseObj = parseJsonMaybe_(ev.case);
    const caseId = caseObj && caseObj.id ? String(caseObj.id) : '';
    if (!caseId) return;

    const rawEventType = firstNonEmpty_(ev.event_type);
    const eventType = String(rawEventType || '').trim().toUpperCase();

    const isRelevantType =
      eventType === 'INITIAL CONSULTATION' ||
      eventType === 'DETAINEE VISITATION';

    if (!isRelevantType) return;

    const startValue = firstNonEmpty_(ev.start);
    if (!startValue) return;

    const currentDate = new Date(String(startValue).replace(/^"+|"+$/g, ''));
    if (isNaN(currentDate.getTime())) return;

    if (!out[caseId]) {
      out[caseId] = {
        first_initial_consultation_date: toDateOnlyMaybe_(startValue),
        first_initial_consultation_title: firstNonEmpty_(ev.name),
        first_initial_consultation_event_type: rawEventType || ''
      };
      return;
    }

    const existingDate = new Date(
      String(out[caseId].first_initial_consultation_date).replace(/^"+|"+$/g, '')
    );

    if (isNaN(existingDate.getTime()) || currentDate < existingDate) {
      out[caseId] = {
        first_initial_consultation_date: toDateOnlyMaybe_(startValue),
        first_initial_consultation_title: firstNonEmpty_(ev.name),
        first_initial_consultation_event_type: rawEventType || ''
      };
    }
  });

  return out;
}

function formatFactCaseMasterDateColumns_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.sheets.factCaseMaster);

  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  ['case_opened_date', 'case_updated_at', 'first_initial_consultation_date']
    .forEach(function(name) {
      const col = headers.indexOf(name) + 1;
      if (col > 0) {
        sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
      }
    });
}

function buildLeadMatchesByCaseId_(cases, mycaseLeadsReport, clientsById) {
  const out = {};

  if (!mycaseLeadsReport || !mycaseLeadsReport.length) return out;

  cases.forEach(function (caseRow) {
    const caseId = String(firstNonEmpty_(caseRow.id, caseRow.case_id) || '');
    if (!caseId) return;

    const linkedClientRef = findPreferredCaseClientRef_(caseRow);
    const linkedClient = resolveClientFromRef_(linkedClientRef, clientsById) || {};

    const caseOpenedDate = toDateOnlyMaybe_(firstNonEmpty_(caseRow.opened_date, caseRow.case_opened_date));
    const caseClientName = normalizeText_(firstNonEmpty_(linkedClient.full_name, buildFullName_(linkedClient), caseRow.name));
    const caseClientEmail = normalizeText_(firstNonEmpty_(linkedClient.email));

    let bestMatch = null;
    let bestScore = 0;

    mycaseLeadsReport.forEach(function (leadRow) {
      const leadConversionDate = toDateOnlyMaybe_(firstNonEmpty_(leadRow['Conversion date'], leadRow.conversion_date));
      const leadName = normalizeText_(firstNonEmpty_(leadRow['Lead name'], leadRow.lead_name));
      const leadEmail = normalizeText_(firstNonEmpty_(leadRow['Email'], leadRow.email, leadRow.lead_email));

      let score = 0;

      if (caseOpenedDate && leadConversionDate && caseOpenedDate === leadConversionDate) score += 3;
      if (caseClientEmail && leadEmail && caseClientEmail === leadEmail) score += 4;
      if (caseClientName && leadName && caseClientName === leadName) score += 2;

      if (score > bestScore) {
        bestScore = score;
        bestMatch = {
          lead_name: firstNonEmpty_(leadRow['Lead name'], leadRow.lead_name),
          lead_email: firstNonEmpty_(leadRow['Email'], leadRow.email, leadRow.lead_email),
          referral_source: firstNonEmpty_(leadRow['Referral source'], leadRow.referral_source),
          conversion_date: firstNonEmpty_(leadRow['Conversion date'], leadRow.conversion_date),
          match_method: buildLeadMatchMethod_(caseOpenedDate, leadConversionDate, caseClientEmail, leadEmail, caseClientName, leadName),
          match_score: score
        };
      }
    });

    if (bestMatch && bestScore >= 5) {
      out[caseId] = bestMatch;
    }
  });

  return out;
}

function buildLeadMatchMethod_(caseOpenedDate, leadConversionDate, caseClientEmail, leadEmail, caseClientName, leadName) {
  const parts = [];
  if (caseOpenedDate && leadConversionDate && caseOpenedDate === leadConversionDate) parts.push('conversion_date');
  if (caseClientEmail && leadEmail && caseClientEmail === leadEmail) parts.push('email');
  if (caseClientName && leadName && caseClientName === leadName) parts.push('name');
  return parts.join('+');
}

function normalizeText_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizePhone_(value) {
  const digits = String(value || '').replace(/\D+/g, '');

  if (!digits) return '';

  // Si tiene más de 10 dígitos (ej: +1), tomar los últimos 10
  if (digits.length > 10) {
    return digits.slice(-10);
  }

  return digits;
}

function classifyLeadType_(leadMatch, consultDateRaw, caseOpenedRaw) {
 
  const consultDate = toDateOnlyMaybe_(consultDateRaw);
  const caseOpened = toDateOnlyMaybe_(caseOpenedRaw);

  if (!consultDate) return 'Existing Client';

  if (consultDate <= caseOpened) return 'New Lead';

  return '';
}

/************************************************************
 * FOR NESTED VALUES
 ************************************************************/

function stringifyIdsDeep_(value) {
  if (Array.isArray(value)) {
    return value.map(stringifyIdsDeep_);
  }

  if (value && typeof value === 'object') {
    const out = {};
    Object.keys(value).forEach(function(key) {
      const v = value[key];

      if (key === 'id' || key.endsWith('_id')) {
        out[key] = v !== null && v !== undefined ? String(v) : '';
      } else {
        out[key] = stringifyIdsDeep_(v);
      }
    });
    return out;
  }

  return value;
}
/**NORMALIZE STRING TO OBJECT DATE*/

function toDateMaybe_(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return value;
  }

  const text = String(value).trim().replace(/^"+|"+$/g, '');
  const d = new Date(value);
  if (isNaN(d.getTime())) {
    return value;
  }

  return d;
}

function toDateOnlyMaybe_(value) {
  if (!value) return '';

  const text = String(value).trim().replace(/^"+|"+$/g, '');
  const d = new Date(text);

  if (isNaN(d.getTime())) {
    return '';
  }

  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function normalizeReferralSource_(leadReferralSource, leadType) {
  const referral = String(leadReferralSource || '').trim();
  const type = String(leadType || '').trim();

  if (!referral && type === 'Existing Client') {
    return 'Existing Client';
  }

  return referral;
}
/************************************************************
 * HTTP HELPERS
 ************************************************************/
function myCaseGetFullResponse_(path, queryParams) {
  const service = getMyCaseService_();

  if (!service.hasAccess()) {
    throw new Error('Primero debes autorizar la conexión ejecutando beginAuth().');
  }

  let url = `${MYCASE_API_BASE}${path}`;

  if (queryParams) {
    const qs = Object.keys(queryParams)
      .filter(k => queryParams[k] !== null && queryParams[k] !== undefined && queryParams[k] !== '')
      .map(k => `${encodeURIComponent(k)}=${encodeURIComponent(queryParams[k])}`)
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

function myCaseGet_(path, queryParams) {
  const response = myCaseGetFullResponse_(path, queryParams);
  return JSON.parse(response.getContentText());
}

function extractNextPageTokenFromLink_(linkHeader) {
  if (!linkHeader) return null;
  const match = linkHeader.match(/[?&]page_token=([^&>]+)/);
  return match ? decodeURIComponent(match[1]) : null;
}

function getAllPaginated_(path, baseParams) {
  let allRecords = [];
  let pageToken = null;
  let keepGoing = true;

  while (keepGoing) {
    const params = Object.assign({}, baseParams || {});
    if (pageToken) {
      params.page_token = pageToken;
    }

    const response = myCaseGetFullResponse_(path, params);
    const parsed = JSON.parse(response.getContentText());
    const records = Array.isArray(parsed) ? parsed : (parsed.data || []);

    allRecords = allRecords.concat(records);

    const headers = response.getHeaders();
    const linkHeader = headers.Link || headers.link || '';
    pageToken = extractNextPageTokenFromLink_(linkHeader);
    keepGoing = !!pageToken;
  }

  return allRecords;
}

/************************************************************
 * GENERIC HELPERS
 ************************************************************/
function safeGet_(obj, path, defaultValue) {
  if (!obj || !path) return defaultValue;
  const parts = path.split('.');
  let current = obj;

  for (let i = 0; i < parts.length; i++) {
    if (current === null || current === undefined || !(parts[i] in current)) {
      return defaultValue;
    }
    current = current[parts[i]];
  }

  return current === undefined ? defaultValue : current;
}

function asJson_(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object') return JSON.stringify(value);
  return String(value);
}

function cleanScalar_(value) {
  if (value === null || value === undefined) return '';
  return value;
}

function uniqueByKey_(rows, key) {
  const map = {};
  rows.forEach(r => {
    map[String(r[key])] = r;
  });
  return Object.keys(map).map(k => map[k]);
}

function fullRefreshCaseMaster() {
   const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    Logger.log('Ya hay una ejecución en curso.');
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
    Logger.log('Duración total: ' + ((new Date() - start) / 1000) + ' segundos');

  } catch (error) {
    Logger.log('ERROR en fullRefreshCaseMaster: ' + error.message);
    Logger.log(error.stack);
    throw error;
    } finally {
    lock.releaseLock();
  }
}

function updateLastRefreshTimestamp_() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Menu');

  sheet.getRange('A1').setValue('Última actualización: ' + new Date());
}