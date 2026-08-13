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
    text: String(firstNonEmpty_(payload.text)).trim(),
    questions: payload.questions || {},
    location: String(firstNonEmpty_(payload.location)).trim(),
    specialty: String(firstNonEmpty_(payload.speciality, payload.specialty)).trim(),
    cost: String(firstNonEmpty_(payload.cost)).trim(),
    referral_source: String(firstNonEmpty_(payload.referral_source, 'ElAbogado.com')).trim(),
    practice_area: String(firstNonEmpty_(payload.practice_area, payload.case_type, payload.matter_type)).trim(),
    raw_payload: payload
  };
}

function buildMyCaseLeadPayloadFromElAbogado_(leadPayload) {
  if (!leadPayload.first_name && !leadPayload.last_name) {
    throw new Error('Missing lead name.');
  }

  const myCasePayload = {
    first_name: leadPayload.first_name,
    last_name: leadPayload.last_name,
    email: leadPayload.email,
    cell_phone_number: leadPayload.phone,
    lead_details: buildElAbogadoLeadDetails_(leadPayload),
    referral_source_reference: {
      id: 5020308
    }
  };

  return myCasePayload;
}

function buildElAbogadoLeadDetails_(leadPayload) {
  const questionLines = formatElAbogadoQuestions_(leadPayload.questions);
  const noteParts = [
    'Source: elAbogado.com',
    leadPayload.specialty ? 'Specialty: ' + leadPayload.specialty : '',
    leadPayload.location ? 'Location: ' + leadPayload.location : '',
    leadPayload.cost ? 'Cost: ' + leadPayload.cost : '',
    leadPayload.practice_area ? 'Practice area: ' + leadPayload.practice_area : '',
    leadPayload.text ? 'Text: ' + leadPayload.text : '',
    questionLines ? 'Questions:\n' + questionLines : '',
    leadPayload.notes ? 'Message: ' + leadPayload.notes : ''
  ].filter(Boolean);

  return noteParts.join('\n');
}

function formatElAbogadoQuestions_(questions) {
  if (!questions || typeof questions !== 'object' || Array.isArray(questions)) {
    return '';
  }

  return Object.keys(questions).map(function(questionKey) {
    return questionKey + ': ' + firstNonEmpty_(questions[questionKey]);
  }).filter(Boolean).join('\n');
}

function extractFirstNameFromFullName_(fullName) {
  const parts = String(firstNonEmpty_(fullName)).trim().split(/\s+/).filter(Boolean);
  return parts.length ? parts[0] : '';
}

function extractLastNameFromFullName_(fullName) {
  const parts = String(firstNonEmpty_(fullName)).trim().split(/\s+/).filter(Boolean);
  return parts.length > 1 ? parts.slice(1).join(' ') : '';
}
