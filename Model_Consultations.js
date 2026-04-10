function buildFactConsultations() {
  const events = readSheetAsObjects_(CONFIG.sheets.rawEvents);

  if (!events || !events.length) {
    ensureSheetExists_(CONFIG.sheets.factConsultations);
    writeRowsToSheet_(CONFIG.sheets.factConsultations, []);
    formatFactConsultationsColumns_();
    return;
  }

  ensureSheetExists_(CONFIG.sheets.factConsultations);

  const rows = [];
  const seen = new Set();

  events.forEach(function(ev) {
    const rawEventType = firstNonEmpty_(ev.event_type);
    const eventType = String(rawEventType || '').trim().toUpperCase();

    const isRelevantType =
      eventType === 'INITIAL CONSULTATION' ||
      eventType === 'DETAINEE VISITATION';

    if (!isRelevantType) return;

    const startValue = firstNonEmpty_(ev.start);
    if (!startValue) return;

    const cleanStart = String(startValue).replace(/^"+|"+$/g, '').trim();
    const eventDateObj = new Date(cleanStart);
    if (isNaN(eventDateObj.getTime())) return;

    const caseObj = parseJsonMaybe_(ev.case);
    const caseId = caseObj && caseObj.id ? String(caseObj.id) : '';

    const locationObj = parseJsonMaybe_(ev.location);
    const staffObj = parseJsonMaybe_(ev.staff);

    const eventId = firstNonEmpty_(ev.id, ev.event_id);
    const dedupeKey = eventId
      ? 'ID_' + String(eventId)
      : [
          caseId,
          eventType,
          cleanStart,
          firstNonEmpty_(ev.name)
        ].join('|');

    if (seen.has(dedupeKey)) return;
    seen.add(dedupeKey);

    rows.push({
      event_id: firstNonEmpty_(ev.id, ev.event_id),
      case_id: caseId,

      event_type: eventType,
      consultation_category:
        eventType === 'INITIAL CONSULTATION'
          ? 'Initial Consultation'
          : 'Detainee Visitation',

      event_name: firstNonEmpty_(ev.name, ev.title),

      event_start_raw: cleanStart,
      event_date: Utilities.formatDate(
        eventDateObj,
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      ),
      event_datetime: Utilities.formatDate(
        eventDateObj,
        Session.getScriptTimeZone(),
        'yyyy-MM-dd HH:mm:ss'
      ),
      event_year: Number(
        Utilities.formatDate(eventDateObj, Session.getScriptTimeZone(), 'yyyy')
      ),
      event_month: Number(
        Utilities.formatDate(eventDateObj, Session.getScriptTimeZone(), 'M')
      ),
      event_day: Number(
        Utilities.formatDate(eventDateObj, Session.getScriptTimeZone(), 'd')
      ),
      event_hour: Number(
        Utilities.formatDate(eventDateObj, Session.getScriptTimeZone(), 'H')
      ),
      event_weekday: Utilities.formatDate(
        eventDateObj,
        Session.getScriptTimeZone(),
        'EEEE'
      ),

      office_location_id: firstNonEmpty_(locationObj && locationObj.id),
      staff_id: firstNonEmpty_(staffObj && staffObj.id),

      created_at: toDateOnlyMaybe_(firstNonEmpty_(ev.created_at)),
      updated_at: toDateOnlyMaybe_(firstNonEmpty_(ev.updated_at))
    });
  });

  rows.sort(function(a, b) {
    const dateCompare = String(a.event_datetime).localeCompare(String(b.event_datetime));
    if (dateCompare !== 0) return dateCompare;

    const typeCompare = String(a.event_type).localeCompare(String(b.event_type));
    if (typeCompare !== 0) return typeCompare;

    return String(a.case_id).localeCompare(String(b.case_id));
  });

  writeRowsToSheet_(CONFIG.sheets.factConsultations, rows);
  formatFactConsultationsColumns_();
}

function formatFactConsultationsColumns_() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.sheets.factConsultations);

  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  [
    'event_date',
    'created_at',
    'updated_at'
  ].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
  });

  ['event_datetime'].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    }
  });

  [
    'event_year',
    'event_month',
    'event_day',
    'event_hour',
    'office_location_id',
    'staff_id'
  ].forEach(function(name) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('0');
    }
  });
}
