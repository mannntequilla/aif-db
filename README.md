# MyCase Apps Script Integration

This project syncs raw data from the MyCase API into Google Sheets and builds reporting tables on top of that raw data.

The codebase is intentionally kept flat so it works cleanly with `clasp push` and the Google Apps Script editor.

## Project Purpose

Main goals:

- Pull raw entities from the MyCase API into Google Sheets
- Import the MyCase leads CSV report from Google Drive
- Build modeled reporting tables such as `fact_case_master`
- Keep entrypoints simple from the Apps Script UI

## Current File Structure

### Core Configuration

- [`Config.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Config.js)
  Central config for API base URL, sheet names, and endpoint paths.

- [`Auth.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Auth.js)
  OAuth setup and token access for MyCase.

- [`Api.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Api.js)
  Low-level HTTP helpers for paginated MyCase API requests.

### Shared Helpers

- [`CoreSheets.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/CoreSheets.js)
  Spreadsheet read/write helpers like `writeRowsToSheet_()` and `readSheetAsObjects_()`.

- [`CoreObjects.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/CoreObjects.js)
  Generic object helpers like `firstNonEmpty_()`, `parseJsonMaybe_()`, and indexing utilities.

- [`CoreDates.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/CoreDates.js)
  Date normalization and formatting helpers.

- [`CoreNormalize.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/CoreNormalize.js)
  Shared normalization helpers currently focused on referral-source cleanup.

- [`CaseMasterHelpers.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/CaseMasterHelpers.js)
  Helpers used by `fact_case_master`, including custom field extraction and invoice aggregation.

- [`Helpers.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Helpers.js)
  Temporary legacy-only file. Active helpers have already been moved into focused files above.

### Sync Layer

- [`Sync_Definitions.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Sync_Definitions.js)
  Declarative definition of raw sync resources. Each resource maps:
  - endpoint
  - destination raw sheet
  - transform function

- [`Sync_Transforms.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Sync_Transforms.js)
  Resource-specific transforms before raw write, mainly for `cases` and `clients`.

- [`Sync_Runner.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Sync_Runner.js)
  Generic sync engine that reads a resource definition and writes its raw output.

- [`Sync.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Sync.js)
  Public wrapper functions used from Apps Script, such as `syncCases()`, `syncClients()`, `syncExpenses()`, and `syncCustomFields()`.

### Services

- [`Service_LeadsReport.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Service_LeadsReport.js)
  Imports the latest MyCase leads CSV from Drive into `raw_mycase_leads_report`.

### Models / Reporting Tables

- [`Model_CaseMaster.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Model_CaseMaster.js)
  Builds `fact_case_master`. Combines cases, clients, invoices, events, lead report data, and custom fields.

- [`Model_Consultations.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Model_Consultations.js)
  Builds `fact_consultations`.

- [`Model_Funnel.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Model_Funnel.js)
  Builds the lead funnel by date from the imported leads report.

- [`Model_Staffing.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Model_Staffing.js)
  Builds a case-to-staff summary table.

### Debug / Exploration

- [`Debug_Events.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Debug_Events.js)
  Debug helpers for raw events and consultation matching.

- [`Debug_Leads.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Debug_Leads.js)
  Debug helpers for leads, spreadsheet headers, token inspection, and raw expenses profiling.

### Entry Points

- [`Main.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Main.js)
  Main operational entrypoints for manual runs and triggers.

Common public functions in `Main.js`:

- `syncAllRaw()`
- `syncCaseMasterInputs()`
- `fullRefreshCaseMaster()`
- `refreshMyCaseLeadsReport()`
- `exploreExpensesRaw()`
- `resetAutoRefreshTrigger()`

## Main Data Flow

### Case Master Flow

1. `syncCaseMasterInputs()`
2. `syncCases()`, `syncClients()`, `syncInvoices()`, `syncEvents()`, `syncCustomFields()`
3. `importLatestMyCaseLeadsReportFromDrive()`
4. `buildFactCaseMaster()`

### Full Refresh Flow

`fullRefreshCaseMaster()` does:

1. Sync core raw inputs
2. Import latest leads report from Drive
3. Rebuild `fact_case_master`
4. Update refresh timestamp in the `Menu` sheet

### Expenses Exploration Flow

`exploreExpensesRaw()` does:

1. `syncExpenses()`
2. `profileExpensesRaw_()`

This is currently a test/exploration flow to inspect whether raw expenses can support a fixed-expenses report.

## Raw Sheets

Current raw sheets include:

- `raw_cases`
- `raw_clients`
- `raw_leads`
- `raw_invoices`
- `raw_expenses`
- `raw_events`
- `raw_roles`
- `raw_calls`
- `raw_tasks`
- `raw_staff`
- `raw_custom_fields`
- `raw_mycase_leads_report`

## Modeled Sheets

Current modeled/report sheets include:

- `fact_case_master`
- `fact_consultations`
- `funnel_leads_by_date`
- `case_staff_summary`
- `debug_expenses_profile`

## Custom Fields

Custom fields are handled in two layers:

1. `syncCustomFields()` loads the `/custom_fields` endpoint into `raw_custom_fields`
2. `buildFactCaseMaster()` resolves a custom field by name and then extracts its value from each case row's `custom_field_values`

Current implemented example:

- `Retainer`

## How To Sync Local Code To Apps Script

Run from this folder in `cmd` or PowerShell:

```bash
git add .
git commit -m "Your commit message"
clasp push
```

To verify what `clasp` sees:

```bash
clasp show-file-status
```

## Recommended Testing Order

For `Retainer` custom field changes:

1. `syncCustomFields()`
2. `syncCaseMasterInputs()`
3. `buildFactCaseMaster()`

For expenses exploration:

1. `exploreExpensesRaw()`

For a complete refresh:

1. `fullRefreshCaseMaster()`

## Maintenance Rules

- Keep the project flat at the filesystem root for reliable `clasp push`
- Put entrypoints in `Main.js` or wrapper functions in `Sync.js`
- Put reusable logic in focused helper/model/service files
- Avoid reintroducing large mixed-purpose helper files
- Prefer declarative sync additions through `Sync_Definitions.js`

## Adding A New Raw Resource

To add a new MyCase endpoint:

1. Add the endpoint and target sheet in [`Config.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Config.js)
2. Add a resource definition in [`Sync_Definitions.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Sync_Definitions.js)
3. Add a public wrapper in [`Sync.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Sync.js)
4. Optionally add it to `syncAllRaw()` or another flow in [`Main.js`](/C:/Users/valer/Aguado_Automations/02_mycase_integrations/mycase_appscript/Main.js)

## Current Notes

- `Helpers.js` is still present only as a temporary legacy container
- The project was recently flattened specifically for Apps Script compatibility
- Expenses are currently exploratory and not yet used in final reporting
