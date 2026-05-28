function doGet(e) {
  return handleProxyRequest_(e, 'GET');
}

function doPost(e) {
  return handleProxyRequest_(e, 'POST');
}

function handleProxyRequest_(e, method) {
  try {
    const action = getProxyActionOrBlank_(e);

    if (action === 'elabogadoWebhook' || isElAbogadoWebhookRequest_(e)) {
      return jsonOk_(handleElAbogadoWebhook_(e));
    }

    validateProxyRequest_(e);

    if (!action) {
      throw new Error('Missing action.');
    }

    if (action === 'health') {
      return jsonOk_(handleHealth_());
    }

    if (action === 'getAccessToken') {
      return jsonOk_(handleGetAccessToken_());
    }

    return jsonError_('Unsupported action: ' + action, 400);
  } catch (error) {
    Logger.log('Proxy error [' + method + ']: ' + (error && error.stack ? error.stack : error));
    return jsonError_(error && error.message ? error.message : String(error), 500);
  }
}

function getProxyActionOrBlank_(e) {
  const payload = getProxyPayload_(e);
  const action = firstNonEmpty_(
    payload.action,
    e && e.parameter ? e.parameter.action : '',
    e && e.parameters && e.parameters.action && e.parameters.action.length ? e.parameters.action[0] : ''
  );

  return action ? String(action).trim() : '';
}

function isElAbogadoWebhookRequest_(e) {
  const payload = getProxyPayload_(e);
  const secret = String(
    firstNonEmpty_(
      payload.secret,
      payload.webhook_secret,
      e && e.parameter ? e.parameter.secret : '',
      e && e.parameter ? e.parameter.webhook_secret : ''
    )
  ).trim();

  const hasLeadName = !!String(
    firstNonEmpty_(
      payload.first_name,
      payload.firstName,
      payload.full_name,
      payload.fullName,
      payload.name
    )
  ).trim();

  return !!(secret && hasLeadName);
}
