function validateProxyRequest_(e) {
  const expectedApiKey = PropertiesService.getScriptProperties().getProperty('INTERNAL_PROXY_API_KEY');

  if (!expectedApiKey) {
    throw new Error('Missing INTERNAL_PROXY_API_KEY in Script Properties.');
  }

  const providedApiKey = getProxyApiKey_(e);
  if (!providedApiKey) {
    throw new Error('Missing API key.');
  }

  if (providedApiKey !== expectedApiKey) {
    throw new Error('Invalid API key.');
  }
}

function getProxyAction_(e) {
  const payload = getProxyPayload_(e);
  const action = firstNonEmpty_(
    payload.action,
    e && e.parameter ? e.parameter.action : '',
    e && e.parameters && e.parameters.action && e.parameters.action.length ? e.parameters.action[0] : ''
  );

  if (!action) {
    throw new Error('Missing action.');
  }

  return String(action).trim();
}

function getProxyApiKey_(e) {
  const payload = getProxyPayload_(e);
  return firstNonEmpty_(
    payload.apiKey,
    payload.api_key,
    e && e.parameter ? e.parameter.apiKey : '',
    e && e.parameter ? e.parameter.api_key : ''
  );
}

function getProxyPayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return {};
  }

  try {
    return JSON.parse(e.postData.contents);
  } catch (error) {
    throw new Error('Invalid JSON payload.');
  }
}

function jsonOk_(data) {
  return jsonResponse_({
    ok: true,
    data: data,
    generatedAt: new Date().toISOString()
  });
}

function jsonError_(message, status) {
  return jsonResponse_({
    ok: false,
    error: message,
    status: status || 500,
    generatedAt: new Date().toISOString()
  });
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
