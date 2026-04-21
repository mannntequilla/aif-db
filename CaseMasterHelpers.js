function aggregateInvoicesByCaseId_(invoices) {
  const out = {};

  invoices.forEach(function(inv) {
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
    const preferred = candidates.find(function(c) {
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
  const sheet = getSpreadsheet_()
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

function getCustomFieldIdByName_(customFields, fieldName, parentType) {
  if (!customFields || !customFields.length) return '';

  const normalizedFieldName = normalizeText_(fieldName);
  const normalizedParentType = normalizeText_(parentType || '');

  const match = customFields.find(function(customField) {
    const currentName = normalizeText_(customField.name);
    const currentParentType = normalizeText_(customField.parent_type);

    if (currentName !== normalizedFieldName) return false;
    if (normalizedParentType && currentParentType !== normalizedParentType) return false;

    return true;
  });

  return match ? String(firstNonEmpty_(match.id)) : '';
}

function getCaseCustomFieldValueById_(caseRow, customFieldId) {
  if (!customFieldId) return '';

  const customFieldValues = parseJsonMaybe_(caseRow.custom_field_values) || [];
  if (!Array.isArray(customFieldValues)) return '';

  const match = customFieldValues.find(function(customFieldValueRow) {
    const customField = customFieldValueRow.custom_field || {};
    return String(firstNonEmpty_(customField.id)) === String(customFieldId);
  });

  return match ? firstNonEmpty_(match.value) : '';
}

function classifyLeadType_(leadMatch, consultDateRaw, caseOpenedRaw) {
  const consultDate = toDateOnlyMaybe_(consultDateRaw);
  const caseOpened = toDateOnlyMaybe_(caseOpenedRaw);

  if (!consultDate) return 'Existing Client';

  if (consultDate <= caseOpened) return 'New Lead';

  return '';
}

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
