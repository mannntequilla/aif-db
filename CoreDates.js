function parseDateOnlyLocal_(value) {
  const text = String(value || '').trim().replace(/^"+|"+$/g, '');
  const match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!match) return null;

  const year = Number(match[1]);
  const monthIndex = Number(match[2]) - 1;
  const day = Number(match[3]);
  const localDate = new Date(year, monthIndex, day);

  if (
    localDate.getFullYear() !== year ||
    localDate.getMonth() !== monthIndex ||
    localDate.getDate() !== day
  ) {
    return null;
  }

  return localDate;
}

function formatDateOnlyForSheet_(value) {
  if (!value) return '';

  const d = parseDateOnlyLocal_(value) || new Date(value);
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
  const d = parseDateOnlyLocal_(text) || new Date(text);

  if (isNaN(d.getTime())) {
    return '';
  }

  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
