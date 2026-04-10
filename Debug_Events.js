function debugInitialConsultEvents() {
  const events = readSheetAsObjects_(CONFIG.sheets.rawEvents);

  events.slice(0, 30).forEach(function(ev, i) {
    Logger.log('--- EVENT ' + i + ' ---');
    Logger.log('id=' + ev.id);
    Logger.log('case_id=' + ev.case_id);
    Logger.log('case=' + ev.case);
    Logger.log('event_type=' + ev.event_type);
    Logger.log('type=' + ev.type);
    Logger.log('title=' + ev.title);
    Logger.log('name=' + ev.name);
    Logger.log('subject=' + ev.subject);
    Logger.log('description=' + ev.description);
    Logger.log('start_at=' + ev.start_at);
    Logger.log('start_time=' + ev.start_time);
    Logger.log('date=' + ev.date);
  });
}

function debugRawEventsHeaders() {
  const events = readSheetAsObjects_(CONFIG.sheets.rawEvents);
  if (!events.length) {
    Logger.log('No events found');
    return;
  }

  Logger.log(JSON.stringify(Object.keys(events[0])));
}

function debugFirstConsultEventType() {
  const events = readSheetAsObjects_(CONFIG.sheets.rawEvents);
  const firstConsultByCaseId = getFirstInitialConsultationByCaseId_(events);

  const keys = Object.keys(firstConsultByCaseId);
  Logger.log('TOTAL CASES WITH CONSULT EVENTS: ' + keys.length);

  if (keys.length) {
    const sample = firstConsultByCaseId[keys[0]];
    Logger.log(JSON.stringify(sample, null, 2));
  }
}

function debugRawEventsHeadersAndSample() {
  const events = readSheetAsObjects_(CONFIG.sheets.rawEvents);

  Logger.log('TOTAL EVENTS: ' + events.length);

  if (!events.length) return;

  Logger.log('EVENT KEYS: ' + JSON.stringify(Object.keys(events[0]), null, 2));
  Logger.log('FIRST EVENT SAMPLE: ' + JSON.stringify(events[0], null, 2));
}

function debugFirstConsultObjects() {
  const events = readSheetAsObjects_(CONFIG.sheets.rawEvents);
  const firstConsultByCaseId = getFirstInitialConsultationByCaseId_(events);

  const keys = Object.keys(firstConsultByCaseId);
  Logger.log('TOTAL MATCHED CASES: ' + keys.length);

  keys.slice(0, 10).forEach(function(key) {
    Logger.log(key + ' => ' + JSON.stringify(firstConsultByCaseId[key], null, 2));
  });
}
