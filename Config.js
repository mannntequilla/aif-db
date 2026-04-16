const MYCASE_AUTH_URL = 'https://auth.mycase.com/login_sessions/new';
const MYCASE_TOKEN_URL = 'https://auth.mycase.com/tokens';

const CONFIG = {
  api: {
    baseUrl: 'https://external-integrations.mycase.com/v1',
    pageSize: 100,
    maxRetries: 5,
    timezone: Session.getScriptTimeZone(),
  },

  sheets: {
    rawCases: 'raw_cases',
    rawClients: 'raw_clients',
    rawLeads: 'raw_leads',
    rawInvoices: 'raw_invoices',
    rawExpenses: 'raw_expenses',
    rawEvents: 'raw_events',
    rawRoles: 'raw_roles',
    rawCalls: 'raw_calls',
    rawTasks: 'raw_tasks',
    rawStaff: 'raw_staff',
    rawCustomFields: 'raw_custom_fields',
    factCaseMaster: 'fact_case_master',
    factCaseProfitability: 'fact_case_profitability',
    factConsultations: 'fact_consultations',
    factLeads: 'fact_leads',
    taskReport: 'task_report',
    rawMyCaseLeadsReport: 'raw_mycase_leads_report',
    funnelLeadsByDate: 'funnel_leads_by_date',
  },

  endpoints: {
    cases: '/cases',
    clients: '/clients',
    leads: '/leads',
    calls: '/calls/',
    tasks: '/tasks/',
    invoices: '/invoices',
    expenses: '/expenses',
    events: '/events',
    roles: '/case_roles',
    staff: '/staff',
    customFields: '/custom_fields',
  }
};
