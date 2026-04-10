function formatDateOnlyForSheet_(value) {
  if (!value) return '';

  const d = new Date(value);
  if (isNaN(d.getTime())) return '';

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}

function toDateMaybe_(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return value;
  }

  const d = new Date(value);
  if (isNaN(d.getTime())) {
    return value;
  }

  return d;
}

function toDateOnlyMaybe_(value) {
  if (!value) return '';

  const text = String(value).trim().replace(/^"+|"+$/g, '');
  const d = new Date(text);

  if (isNaN(d.getTime())) {
    return '';
  }

  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
