function buildCaseStaffTable (){
  const casesSheet = readSheetAsObjects_(CONFIG.sheets.rawCases);
  const staffSheet = readSheetAsObjects_(CONFIG.sheets.rawStaff);
  const staffById = {};

  staffSheet.forEach(s => {
    const id = String(s.id).trim();
    staffById[id] = s;
  });

  const output = [];

  casesSheet.forEach(c => {
    const caseId = c.id;
    const caseName = c.name;
    const caseStaff = c.staff;

    let assignedStaffNames = [];
    let assignedStaffIds = [];

     if (caseStaff) {
      const parsedStaff = parseJsonMaybe_(caseStaff);
  
    if (Array.isArray(parsedStaff)) {
        parsedStaff.forEach(member => {
          const staffId = String(member.id).trim();
          const staffMatch = staffById[staffId];

          const fullName = staffMatch
            ? [staffMatch.first_name, staffMatch.last_name].filter(Boolean).join(' ')
            : `ID ${staffId}`;

          assignedStaffNames.push(fullName);
          assignedStaffIds.push(staffId);
        });
      }
    }

      output.push({
      case_id: caseId,
      case_name: caseName,
      assigned_staff_names: assignedStaffNames.join(', '),
      assigned_staff_ids: assignedStaffIds.join(', '),
      has_staff_assigned: assignedStaffNames.length > 0 ? 'Yes' : 'No'
    });
  });

  writeRowsToSheet_('case_staff_summary', output);
}

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

    // Dedupe básico por id; si no existe id, usa combinación estable
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

  // Orden recomendado: fecha ascendente, luego tipo, luego case
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

function ensureSheetExists_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  return sheet;
}
