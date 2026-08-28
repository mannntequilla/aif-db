/**
 * Task reporting fact at task-assignee grain. A task assigned to more than one
 * staff member produces one row per assignment so workload can be analyzed by
 * person. Use COUNT_DISTINCT(task_key) for firm-level task totals.
 */
function buildFactTask() {
  const tasks = readSheetAsObjectsIfExists_(CONFIG.sheets.rawTasks);
  const staff = readSheetAsObjectsIfExists_(CONFIG.sheets.rawStaff);
  const staffById = indexBy_(staff, 'id');
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const rows = [];

  tasks.forEach(function(taskRow) {
    const taskKey = String(firstNonEmpty_(taskRow.id, taskRow.task_id)).trim();
    if (!taskKey) return;

    const taskName = firstNonEmpty_(taskRow.name, taskRow.task_name);
    const createdDate = toDateOnlyMaybe_(firstNonEmpty_(taskRow.created_at, taskRow.task_created_at));
    const dueDate = toDateOnlyMaybe_(firstNonEmpty_(taskRow.due_date, taskRow.due_at));
    const completedDate = toDateOnlyMaybe_(firstNonEmpty_(taskRow.completed_at, taskRow.completed_date));
    const updatedDate = toDateOnlyMaybe_(firstNonEmpty_(taskRow.updated_at, taskRow.task_updated_at));
    const isCompleted = normalizeTaskCompleted_(taskRow.completed, completedDate);
    const caseKey = extractTaskCaseKey_(taskRow);
    const assigneeKeys = extractTaskAssigneeKeys_(taskRow);
    const timing = getTaskTiming_(createdDate, dueDate, completedDate, isCompleted, today);
    const assignments = assigneeKeys.length ? assigneeKeys : [''];

    assignments.forEach(function(staffKey) {
      const staffRow = staffKey ? (staffById[staffKey] || {}) : {};
      const staffName = staffKey ? firstNonEmpty_(staffRow.full_name, buildFullName_(staffRow)) : 'Unassigned';

      rows.push({
        task_assignment_key: [taskKey, staffKey || 'unassigned'].join('|'),
        task_key: taskKey,
        case_key: caseKey,
        task_name: taskName,
        task_name_key: normalizeDimensionKey_(taskName),
        task_description: firstNonEmpty_(taskRow.description),
        priority: firstNonEmpty_(taskRow.priority, 'Unspecified'),

        assigned_staff_key: staffKey,
        assigned_staff_name: staffName || ('Unknown staff ' + staffKey),
        assigned_staff_title: firstNonEmpty_(staffRow.title),
        assigned_staff_active: firstNonEmpty_(staffRow.active),
        is_assigned: staffKey ? 1 : 0,
        assigned_staff_count: assigneeKeys.length,
        is_multi_assignee: assigneeKeys.length > 1 ? 1 : 0,

        created_date: createdDate,
        due_date: dueDate,
        completed_date: completedDate,
        updated_date: updatedDate,

        is_completed: isCompleted ? 1 : 0,
        is_pending: isCompleted ? 0 : 1,
        has_due_date: dueDate ? 1 : 0,
        is_overdue: timing.is_overdue,
        is_completed_late: timing.is_completed_late,
        days_to_complete: timing.days_to_complete,
        days_overdue: timing.days_overdue,
        task_timing_status: timing.status,

        task_count: 1
      });
    });
  });

  writeRowsToSheet_(CONFIG.sheets.factTask, rows);
  formatFactTaskColumns_();
}

function extractTaskCaseKey_(taskRow) {
  const taskCase = parseJsonMaybe_(firstNonEmpty_(taskRow.case, '{}')) || {};
  return String(firstNonEmpty_(taskRow.case_id, taskCase.id, taskCase.case_id)).trim();
}

function extractTaskAssigneeKeys_(taskRow) {
  const rawStaff = parseJsonMaybe_(firstNonEmpty_(taskRow.staff, '[]'));
  const assignments = Array.isArray(rawStaff) ? rawStaff : (rawStaff ? [rawStaff] : []);
  const seen = {};

  return assignments.map(function(assignment) {
    return String(firstNonEmpty_(assignment.id, assignment.staff_id)).trim();
  }).filter(function(staffKey) {
    if (!staffKey || seen[staffKey]) return false;
    seen[staffKey] = true;
    return true;
  });
}

function normalizeTaskCompleted_(completedValue, completedDate) {
  if (completedDate) return true;
  return completedValue === true || String(completedValue).toLowerCase() === 'true';
}

function getTaskTiming_(createdDate, dueDate, completedDate, isCompleted, today) {
  const completionReference = isCompleted ? completedDate : today;
  const isCompletedLate = Boolean(isCompleted && dueDate && completedDate && completedDate > dueDate);
  const isOverdue = Boolean(!isCompleted && dueDate && today > dueDate);
  const daysOverdue = dueDate && completionReference && completionReference > dueDate ?
    daysBetweenDates_(dueDate, completionReference) : 0;
  const daysToComplete = isCompleted && createdDate && completedDate ?
    daysBetweenDates_(createdDate, completedDate) : '';
  let status = '';

  if (!dueDate) {
    status = isCompleted ? 'Completed - No Due Date' : 'Pending - No Due Date';
  } else if (!isCompleted) {
    status = isOverdue ? 'Pending - Overdue' : 'Pending - On Time';
  } else {
    status = isCompletedLate ? 'Completed - Late' : 'Completed - On Time';
  }

  return {
    is_overdue: isOverdue ? 1 : 0,
    is_completed_late: isCompletedLate ? 1 : 0,
    days_to_complete: daysToComplete,
    days_overdue: daysOverdue,
    status: status
  };
}

function formatFactTaskColumns_() {
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.sheets.factTask);
  if (!sheet || sheet.getLastRow() < 2) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  ['created_date', 'due_date', 'completed_date', 'updated_date'].forEach(function(name) {
    const column = headers.indexOf(name) + 1;
    if (column > 0) sheet.getRange(2, column, sheet.getLastRow() - 1, 1).setNumberFormat('yyyy-mm-dd');
  });
}
