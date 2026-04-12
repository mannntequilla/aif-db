function handleHealth_() {
  const service = getMyCaseService_();
  const tokenData = getStoredMyCaseTokenInfo_();

  return {
    service: 'mycase-token-broker',
    authorized: service.hasAccess(),
    hasRefreshToken: !!(tokenData && tokenData.refresh_token),
    hasAccessToken: !!(tokenData && tokenData.access_token),
    expiresAt: tokenData && tokenData.expiresAt ? tokenData.expiresAt : null
  };
}

function handleGetAccessToken_() {
  const tokenInfo = getValidAccessTokenInfo_();

  return {
    accessToken: tokenInfo.accessToken,
    expiresAt: tokenInfo.expiresAt,
    tokenType: tokenInfo.tokenType
  };
}

function getValidAccessTokenInfo_() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Could not acquire proxy auth lock.');
  }

  try {
    const accessToken = getAccessToken_();
    const tokenData = getStoredMyCaseTokenInfo_();

    return {
      accessToken: accessToken,
      expiresAt: tokenData && tokenData.expiresAt ? tokenData.expiresAt : null,
      tokenType: tokenData && tokenData.token_type ? tokenData.token_type : 'Bearer'
    };
  } finally {
    lock.releaseLock();
  }
}

function getStoredMyCaseTokenInfo_() {
  const rawToken = PropertiesService.getUserProperties().getProperty('oauth2.mycase');
  if (!rawToken) {
    return null;
  }

  try {
    const tokenData = JSON.parse(rawToken);
    return {
      access_token: tokenData.access_token || '',
      refresh_token: tokenData.refresh_token || '',
      token_type: tokenData.token_type || 'Bearer',
      expiresAt: tokenData.expiresAt || tokenData.expires_at || null
    };
  } catch (error) {
    Logger.log('Unable to parse stored MyCase token: ' + error.message);
    return null;
  }
}
