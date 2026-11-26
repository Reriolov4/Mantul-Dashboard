# Mantul Dashboard — Page 2/3/4 Revision

Date: 25 November 2026
Target branch: `perf/page2-single-pass`
Production deployment: existing Apps Script Web App deployment ID / URL (via GitHub Actions secrets)

## Implemented in this PR

### Page 1
- Business content, KPIs, formulas, charts, filters, and matrix are untouched.
- Only the navigation labels in `Index.html` were changed.

### Page 2
- Existing Page 2 implementation is retained.
- The derived-index rebuild is protected by an explicit confirmation token.
- Normal dashboard refresh no longer calls `analyticsBuildDashboardIndexes()`.

### Page 3 — Outlets
- Dedicated `Page3Service.gs`, `Page3.html`, and `Page3Scripts.html`.
- SOO scope includes every TIM present in the maintained data; no SALESMAN-only restriction.
- SO mode reads all sales types and uses SUB_DIST; excludes `0` and maps AP. GEUBRINA RIZKI to EPM LSE and PT RESTU INDO MEDICA to EPM BAC.
- Both modes retain the visible entity label `SUBDIST`.
- Customer Name search, year/month selection, Subdist, Subbrand, Value/Karton toggle.
- Monthly trend table by Customer Name + Subdist, with selected months, total, and pagination.
- No Target, ACH, TRG BE, or %BE columns.
- Page 3 reads the 18-column `P2_EVENT_INDEX` directly so SOO and SO can be separated with their respective SUBDIST source columns.

### Page 4 — SKU KCH
- Dedicated `Page4Service.gs`, `Page4.html`, and `Page4Scripts.html`.
- Customer and SUBDIST filters only, matching the Page 2 filter pattern.
- SOO-only YTD aggregation from `P2_EVENT_INDEX`, with distributor identity read from `SUBDIST`.
- SKU KCH resolved from the `SKU` field; SKU master mapping uses `SKU.PRODUCT_DESC`.
- Ambiguous product attributes are labeled `MULTI` instead of assigning an arbitrary Brand/Subbrand/Category/Entitas.

### Safe index handling
- `AnalyticsIndexAudit.gs` validates existing derived sheets against `P2_EVENT_INDEX` and registers readiness/range metadata in Script Properties only.
- It does not modify any spreadsheet sheet.
- `DashboardRefreshService.gs` runs that audit during refresh and never rebuilds derived sheets.
- `analyticsBuildDashboardIndexes()` now requires the explicit argument `REBUILD_DERIVED_INDEXES`; accidental invocation without it throws before any write.

## Workbook reconciliation completed

Against the supplied workbook `SALES PERFORMANCE DATABASE 2026(1).xlsx`:

- `P2_EVENT_INDEX` contains 127,285 SOO transaction rows.
- `P2_DASHBOARD_FACT` has 88,970 groups; its `SOURCE_ROWS` sum equals 127,285.
- `P2_DASHBOARD_FACT` and `P2_SUBDIST_FACT` totals reconcile to SOO source totals:
  - Value: 52,751,374,743.05
  - Carton: 85,086.79646944
- `P3_OUTLET_MONTHLY` reconciles to the SALESMAN-only source subset:
  - Value: 47,144,202,167.76
  - Carton: 67,548.504802911
- `P2_SUBDIST_FACT` is ordered by Subdist and contains 7 distinct Subdist blocks.
- `RAW_DATA_SALES` has 63,823 SOO transaction rows for 2026 (Jan–Nov).
- SKU master has 81 distinct `PRODUCT_DESC` values; 80 have transaction activity in the supplied workbook. Duplicate/ambiguous master attributes are not silently treated as unique.

## Static checks and limitations

- JavaScript syntax checks passed for the local Page 3/Page 4 scripts and Apps Script service files.
- Workbook row-count and total reconciliation passed.
- This is not a live Apps Script execution test. After GitHub Actions deploys, verify the existing Web App URL, authorize if prompted, and test each page against the live Google Sheet.
- Deployment workflow uses `clasp push --force` and updates the existing deployment ID. A failed Actions run must be fixed before declaring deployment complete.

## Do not run
- Do not run `analyticsBuildDashboardIndexes()` from the Apps Script editor.
- Do not manually rebuild or delete any derived sheet during rollout.
