function doGet(e) {
  return handleProxyRequest_(e, 'GET');
}

function doPost(e) {
  return handleProxyRequest_(e, 'POST');
}

function handleProxyRequest_(e, method) {
  try {
    validateProxyRequest_(e);

    const action = getProxyAction_(e);

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
