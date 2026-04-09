function myCaseGetFullResponse_(endpoint, params = {}) {
  const token = getAccessToken_();

  const queryString = Object.keys(params)
    .map(function(key) {
      return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
    })
    .join('&');

  const url = CONFIG.api.baseUrl + endpoint + (queryString ? '?' + queryString : '');

  const options = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/json'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();

  if (code >= 400) {
    throw new Error('MyCase API error ' + code + ': ' + response.getContentText());
  }

  return response;
}

function apiGetAllPages_(endpoint, params = {}) {
  let allData = [];
  let pageToken = null;
  let keepGoing = true;

  while (keepGoing) {
    const query = Object.assign({}, params, {
      page_size: CONFIG.api.pageSize
    });

    if (pageToken) {
      query.page_token = pageToken;
    }

    const response = myCaseGetFullResponse_(endpoint, query);
    const data = JSON.parse(response.getContentText());
    const items = Array.isArray(data) ? data : (data.data || []);

    allData = allData.concat(items);

    const headers = response.getHeaders();
    const linkHeader = headers.Link || headers.link || '';
    pageToken = extractNextPageTokenFromLink_(linkHeader);
    keepGoing = !!pageToken;
  }

  return allData;
}

function extractNextPageTokenFromLink_(linkHeader) {
  if (!linkHeader) return null;

  const match = linkHeader.match(/page_token=([^&>]+)/);
  return match ? match[1] : null;
}