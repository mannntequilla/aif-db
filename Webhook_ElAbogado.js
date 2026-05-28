function handleElAbogadoWebhook_(e) {
  const payload = getProxyPayload_(e);
  validateElAbogadoWebhook_(e, payload);

  const normalizedPayload = normalizeElAbogadoWebhookPayload_(payload);
  const myCaseLeadPayload = buildMyCaseLeadPayloadFromElAbogado_(normalizedPayload);
  const apiResponse = myCasePost_(CONFIG.endpoints.leads, myCaseLeadPayload);

  return {
    source: 'elabogado.com',
    receivedLead: normalizedPayload,
    myCasePayload: myCaseLeadPayload,
    myCaseResponse: apiResponse
  };
}

function validateElAbogadoWebhook_(e, payload) {
  const expectedSecret = PropertiesService.getScriptProperties().getProperty('ELABOGADO_WEBHOOK_SECRET');

  if (!expectedSecret) {
    throw new Error('Missing ELABOGADO_WEBHOOK_SECRET in Script Properties.');
  }

  const providedSecret = String(
    firstNonEmpty_(
      payload.secret,
      payload.webhook_secret,
      payload.apiKey,
      payload.api_key,
      e && e.parameter ? e.parameter.secret : '',
      e && e.parameter ? e.parameter.webhook_secret : ''
    )
  ).trim();

  if (!providedSecret) {
    throw new Error('Missing webhook secret.');
  }

  if (providedSecret !== expectedSecret) {
    throw new Error('Invalid webhook secret.');
  }
}

function normalizeElAbogadoWebhookPayload_(payload) {
  const firstName = String(
    firstNonEmpty_(
      payload.first_name,
      payload.firstName,
      extractFirstNameFromFullName_(payload.full_name || payload.fullName || payload.name)
    )
  ).trim();

  const lastName = String(
    firstNonEmpty_(
      payload.last_name,
      payload.lastName,
      extractLastNameFromFullName_(payload.full_name || payload.fullName || payload.name)
    )
  ).trim();

  return {
    first_name: firstName,
    last_name: lastName,
    email: String(firstNonEmpty_(payload.email, payload.email_address)).trim(),
    phone: String(firstNonEmpty_(payload.phone, payload.phone_number, payload.mobile)).trim(),
    notes: String(firstNonEmpty_(payload.notes, payload.message, payload.description)).trim(),
    referral_source: String(firstNonEmpty_(payload.referral_source, 'ElAbogado.com')).trim(),
    practice_area: String(firstNonEmpty_(payload.practice_area, payload.case_type, payload.matter_type)).trim(),
    raw_payload: payload
  };
}

function buildMyCaseLeadPayloadFromElAbogado_(leadPayload) {
  if (!leadPayload.first_name && !leadPayload.last_name) {
    throw new Error('Missing lead name.');
  }

  return {
    first_name: leadPayload.first_name,
    last_name: leadPayload.last_name,
    email: leadPayload.email,
    phone_number: leadPayload.phone,
    referral_source: leadPayload.referral_source,
    practice_area: leadPayload.practice_area,
    notes: buildElAbogadoLeadNotes_(leadPayload)
  };
}

function buildElAbogadoLeadNotes_(leadPayload) {
  const noteParts = [
    'Source: elAbogado.com',
    leadPayload.notes ? 'Message: ' + leadPayload.notes : '',
    'Raw payload: ' + JSON.stringify(leadPayload.raw_payload || {})
  ].filter(Boolean);

  return noteParts.join('\n');
}

function extractFirstNameFromFullName_(fullName) {
  const parts = String(firstNonEmpty_(fullName)).trim().split(/\s+/).filter(Boolean);
  return parts.length ? parts[0] : '';
}

function extractLastNameFromFullName_(fullName) {
  const parts = String(firstNonEmpty_(fullName)).trim().split(/\s+/).filter(Boolean);
  return parts.length > 1 ? parts.slice(1).join(' ') : '';
}
