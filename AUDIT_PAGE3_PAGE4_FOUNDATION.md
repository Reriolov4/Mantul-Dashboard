# Mantul Dashboard — Page 2/3/4 Audit Foundation

Date: 2026-11-25
Base branch: `perf/page2-single-pass`
Related PR: #1 (kept separate)

## Scope of this branch

This is a safety/audit foundation only. It does not claim that Page 3 (Outlets) or Page 4 (SKU KCH) is implemented or production-ready.

### Changes

- `DashboardRefreshService.gs`: `refreshMantulDashboard()` now performs read-only source validation and reports derived-index status. It deliberately does not call `analyticsStartInitialBuild()` or `analyticsBuildDashboardIndexes()`.
- `AnalyticsIndexAudit.gs`: adds `analyticsAuditDerivedIndexesReadOnly()`, which reports sheet headers, row counts, stored index metadata, and consistency checks without modifying data or script properties.

## Workbook audit observations

Based on the supplied workbook `SALES PERFORMANCE DATABASE 2026(1).xlsx`:

- `P1_FACT`: 14,738 non-empty data rows.
- `P2_DASHBOARD_FACT`: 88,970 non-empty data rows.
- `P2_SUBDIST_FACT`: 88,970 non-empty data rows.
- `P3_OUTLET_MONTHLY_BUILD_3794283`: 28,541 non-empty data rows.
- `P3_OUTLET_MONTHLY`: 28,541 non-empty data rows.
- `P2_EVENT_INDEX`: 130,141 non-empty data rows.
- `RAW_DATA_SALES`: 130,141 non-empty data rows.
- `SKU`: 998 non-empty data rows.
- `MAPPING TEAM`: 999 non-empty data rows.

The apparent 577-row difference previously seen between `P2_EVENT_INDEX` and `RAW_DATA_SALES` was due to blank XML rows, not missing non-empty transaction records.

The two Page 2 derived sheets have the same row count and matching aggregate totals in the workbook audit. They are not deleted or rebuilt by this branch.

## Existing Page 3 / Page 4 code gap

The current `Page3Scripts.html` and `Page4Scripts.html` still call `getPage2DashboardData(f)`. They are therefore sharing the Page 2 data endpoint and are not independent implementations of the requested Outlets and SKU KCH pages.

Required next implementation:

- Page 3: Customer Name search, Value/Karton metric, Subdist and Subbrand filters, and a monthly customer trend table without Target/ACH/TRG BE/%BE columns.
- Page 4: Subdist filter and SKU KCH table, with the SKU mapping and transaction aggregation confirmed against the workbook's actual SKU sheet before finalizing the join.

## Safety

- Do not run `analyticsBuildDashboardIndexes()` during this audit phase.
- Do not merge this branch into `perf/page2-single-pass` until Page 3/4 services and tests are completed.
- The repository workflow is configured to deploy when `perf/page2-single-pass` receives a push. This audit branch does not trigger that deployment workflow.
