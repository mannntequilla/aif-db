function indexBy_(rows, key) {
  const out = {};
  rows.forEach(function(row) {
    const value = firstNonEmpty_(row[key]);
    if (value !== '' && value !== null && value !== undefined) {
      out[String(value)] = row;
    }
  });
  return out;
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
  rows.forEach(function(row) {
    map[String(row[key])] = row;
  });
  return Object.keys(map).map(function(mapKey) {
    return map[mapKey];
  });
}
