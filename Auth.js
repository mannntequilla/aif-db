function getMyCaseService_() {
  const props = PropertiesService.getScriptProperties();

  return OAuth2.createService('mycase')
    .setAuthorizationBaseUrl(MYCASE_AUTH_URL)
    .setTokenUrl(MYCASE_TOKEN_URL)
    .setClientId(props.getProperty('MYCASE_CLIENT_ID'))
    .setClientSecret(props.getProperty('MYCASE_CLIENT_SECRET'))
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setParam('response_type', 'code');
}

function logRedirectUri() {
  const service = getMyCaseService_();
  Logger.log('Redirect URI: ' + service.getRedirectUri());
}

function beginAuth() {
  const service = getMyCaseService_();

  if (service.hasAccess()) {
    Logger.log('Ya tienes acceso autorizado.');
    return;
  }

  const authUrl = service.getAuthorizationUrl();
  Logger.log('Abre esta URL en tu navegador: %s', authUrl);
}

function authCallback(request) {
  const service = getMyCaseService_();
  const authorized = service.handleCallback(request);

  if (authorized) {
    return HtmlService.createHtmlOutput('Autorización completada. Puedes cerrar esta pestaña.');
  }

  return HtmlService.createHtmlOutput('Autorización denegada o fallida.');
}

function resetAuth() {
  getMyCaseService_().reset();
}

function getAccessToken_() {
  const service = getMyCaseService_();

  if (!service.hasAccess()) {
    throw new Error('MyCase no está autorizado todavía. Ejecuta beginAuth() primero.');
  }

  return service.getAccessToken();
}