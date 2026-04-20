function buildDimDate() {
  const dateValues = []
    .concat(collectDatesFromRawCases_())
    .concat(collectDatesFromRawEvents_())
    .concat(collectDatesFromRawMyCaseLeadsReport_());

  const validDates = dateValues.filter(function(value) {
    return !!toDateOnlyMaybe_(value);
  }).map(function(value) {
    return toDateOnlyMaybe_(value);
  });

  const fallbackToday = toDateOnlyMaybe_(new Date());
  const minDate = validDates.length ? new Date(Math.min.apply(null, validDates.map(function(d) { return d.getTime(); }))) : fallbackToday;
  const maxDate = validDates.length ? new Date(Math.max.apply(null, validDates.map(function(d) { return d.getTime(); }))) : fallbackToday;

  const startDate = new Date(minDate.getFullYear(), minDate.getMonth(), minDate.getDate());
  const endDate = new Date(maxDate.getFullYear(), maxDate.getMonth(), maxDate.getDate());
  const rows = [];

  for (let current = new Date(startDate); current.getTime() <= endDate.getTime(); current.setDate(current.getDate() + 1)) {
    const rowDate = new Date(current.getFullYear(), current.getMonth(), current.getDate());
    rows.push(buildDimDateRow_(rowDate));
  }

  writeRowsToSheet_(CONFIG.sheets.dimDate, rows);
  formatDimDateColumns_();
}

function collectDatesFromRawCases_() {
  const rows = readSheetAsObjectsIfExists_(CONFIG.sheets.rawCases);
  return rows.map(function(row) {
    return firstNonEmpty_(row.opened_date, row.case_opened_date, row.updated_at, row.case_updated_at);
  });
}

function collectDatesFromRawEvents_() {
  const rows = readSheetAsObjectsIfExists_(CONFIG.sheets.rawEvents);
  return rows.map(function(row) {
    return firstNonEmpty_(row.start, row.start_at, row.start_time, row.date, row.created_at);
  });
}

function collectDatesFromRawMyCaseLeadsReport_() {
  const rows = readSheetAsObjectsIfExists_(CONFIG.sheets.rawMyCaseLeadsReport);
  return rows.map(function(row) {
    return firstNonEmpty_(row['Date added'], row['Conversion date']);
  });
}

function buildDimDateRow_(dateValue) {
  const year = Number(Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy'));
  const month = Number(Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'M'));
  const day = Number(Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'd'));
  const quarter = Math.floor((month - 1) / 3) + 1;
  const monthName = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'MMMM');
  const weekdayName = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'EEEE');
  const yearMonth = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM');
  const yearQuarter = year + '-Q' + quarter;
  const isWeekend = weekdayName === 'Saturday' || weekdayName === 'Sunday' ? 'Yes' : 'No';

  return {
    date: Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    year: year,
    quarter: quarter,
    year_quarter: yearQuarter,
    month: month,
    month_name: monthName,
    year_month: yearMonth,
    day: day,
    weekday_name: weekdayName,
    is_weekend: isWeekend
  };
}

function formatDimDateColumns_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.dimDate);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  ['date'].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
  });

  ['year', 'quarter', 'month', 'day'].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('0');
    }
  });
}
