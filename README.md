# StockFlow Ledger - Excel Side (v2.0 Enterprise Update)

## What was missing before, and what's fixed now

| Gap (from earlier review) | File that fixes it |
|---|---|
| `no-cors` fake-success bug (errors were invisible) | `utils/secureRequest.ts` |
| No security / request signing | `utils/secureRequest.ts` (HMAC-SHA256) |
| No offline retry queue | `utils/syncQueue.ts` |
| No conflict detection between Excel and Sheets edits | `utils/conflictResolver.ts` |
| No real audit trail (just one status cell) | `utils/auditLogger.ts` + `excelController.ts` (hidden `AuditLog` sheet) |
| Single-click "approval" with no second check | `utils/approvalWorkflow.ts` + `components/ApprovalPanel.tsx` |
| No trend visibility, no lead-time learning | `utils/analyticsEngine.ts` + `components/TrendChart.tsx` |
| All Office.js logic crammed inside App.tsx | `utils/excelController.ts` |
| No visible sync/queue status | `components/StatusIndicator.tsx` |

## New folder structure

```
src/
  types.ts                      shared interfaces
  App.tsx                       main taskpane (now an orchestrator, not a dumping ground)
  utils/
    secureRequest.ts            HMAC-signed fetch wrapper, real error handling
    syncQueue.ts                offline retry queue (localStorage-backed)
    conflictResolver.ts         dual-edit detection + resolution strategy
    auditLogger.ts              local audit trail + CSV export
    approvalWorkflow.ts         draft -> pending_finance -> approved/rejected
    analyticsEngine.ts          trend points + adaptive vendor lead-time
    excelController.ts          all Office.js calls, isolated from UI
  components/
    StatusIndicator.tsx         live sync status dot + queue badge
    ApprovalPanel.tsx           two-step approval UI
    TrendChart.tsx              stock trend chart (recharts)
    AuditLogPanel.tsx           recent activity + CSV export button
```

## New dependencies to install

```bash
npm install recharts
```
(`@mui/material` and `@mui/icons-material` were already in use from v1.)

## Before you deploy

1. Replace `GOOGLE_SCRIPT_API_URL` in `App.tsx` with your real deployed Web App URL.
2. Replace `SHARED_SECRET` in `utils/secureRequest.ts` with a real secret,
   and configure the **same** secret in `Code.gs` on the Sheets side to verify
   the `X-Signature` header (HMAC-SHA256) - this is the missing half of the
   security story and belongs to the Sheets-side build.
3. The Google Apps Script Web App must NOT be called with `mode: "no-cors"`
   anymore - it needs to be deployed so a normal CORS response can be read,
   otherwise `secureRequest.ts` can't detect real failures either.
4. `CURRENT_USER` is currently hardcoded - wire it to `Office.context` user
   info (or your auth system) once available.

## Still outside Excel-side scope (belongs to Sheets/backend build)

- `Code.gs` (doPost handler, HMAC verification, onEdit ledger)
- Multi-warehouse consolidation
- WebSocket/relay layer to replace polling with real push events
- Server-side audit log (Excel side now has its own local + in-workbook log,
  but a central log across both sides still needs backend work)

## Excel-side completion estimate after this update

Core sync, security, reliability, approval chain, audit trail, and analytics
are now implemented. Remaining Excel-side items are the near-instant/push
sync upgrade (currently polling-based, which is an acceptable v1 pattern)
and wiring `CURRENT_USER` to real Office identity.




# StockFlow Ledger - Google Sheets Side (Backend + Sidebar)

## 1. Spreadsheet setup

Create a Google Sheet named **StockFlow Warehouse Ledger** with one tab
named exactly **Inventory** (case-sensitive - the backend looks for this
name). Add these headers in row 1:

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| Item Code | Physical Stock | Expected Cargo | Lead Time | Daily Sales Rate | Last Updated | Status |

Add a few sample rows under the header so there's data to test against,
e.g. `ITEM-001, 40, 20, 5, 10, (leave blank), (leave blank)`.

Do NOT create the `AuditLog` sheet manually - `Code.gs` creates and hides
it automatically the first time it's needed.

## 2. Apps Script project setup

1. In the Sheet: **Extensions → Apps Script**. Name the project
   **StockFlow-Ledger-Backend**.
2. Delete the default empty `Code.gs` content, then create/paste these
   files exactly as named (use the `+` next to "Files" to add each one):
   - `Config.gs`
   - `Security.gs`
   - `AuditLogger.gs`
   - `Triggers.gs`
   - `Code.gs`
   - `Sidebar.html` (use **HTML file** type, not a `.gs` file)
3. Open **Project Settings** (gear icon) → under "Show appsscript.json
   manifest file in editor", enable it. Then paste the contents of our
   `appsscript.json` into that file.

## 3. Set your secrets (run once)

In the Apps Script editor, select `setupScriptProperties` from the
function dropdown at the top and click **Run**. Google will ask for
authorization the first time - accept it.

Before running, edit the two placeholder values inside `Config.gs`:
```javascript
SHARED_SECRET: "REPLACE_WITH_THE_SAME_SECRET_USED_IN_EXCEL_SIDE",
API_KEY: "REPLACE_WITH_A_SEPARATE_READ_ACCESS_KEY_FOR_GET_REQUESTS",
```
- `SHARED_SECRET` must be **character-for-character identical** to the
  `SHARED_SECRET` constant in the Excel side's `secureRequest.ts`.
- `API_KEY` must match the `API_KEY` constant in that same Excel file.

## 4. Deploy as a Web App (this gives you the URL Excel needs)

1. Click **Deploy → New deployment**.
2. Type: **Web app**.
3. Execute as: **Me**.
4. Who has access: **Anyone** (or "Anyone within [your org]" if you're on
   Google Workspace and want it restricted).
5. Click **Deploy**, authorize again if asked, then copy the **Web app URL**.
6. Paste that URL into `GOOGLE_SCRIPT_API_URL` in the Excel side's `App.tsx`.

Whenever you edit any `.gs` file, you must create a **new deployment
version** (Deploy → Manage deployments → Edit → New version) for the
Excel side to see the changes - saving the file alone is not enough.

## 5. Reload the sheet and open the sidebar

Refresh the Google Sheet tab. A new menu **StockFlow Ledger** appears at
the top. Click **StockFlow Ledger → Open Control Panel** to see the
sidebar (connection status, item lookup, recent activity).

## What the sidebar shows

- A live-pulsing status dot while it checks the connection
- Which spreadsheet it's connected to and how many items are tracked
- A search box to look up any item code and see its stockout risk
- The 8 most recent audit events (warehouse edits + Excel writebacks)

The sidebar is read-only by design - actual data entry happens directly
in the `Inventory` sheet cells (the `onEdit` trigger picks up every
change automatically), which matches how a warehouse team already works.

## Important correction vs. the original Excel-side draft

Apps Script Web Apps cannot read custom HTTP headers. This backend
expects the HMAC signature as a `signature` field inside the JSON body
(for `doPost`) and the API key as a `?key=` query parameter (for
`doGet`) - the Excel side's `secureRequest.ts` was updated to match.

## Rebuilding the sidebar UI (if you want to change the design)

The compiled `Sidebar.html` was built from real React + MUI source in
`sidebar-source/`. To modify it:

```bash
cd sidebar-source
npm install
npm run build          # outputs sidebar-source/dist/bundle.js
```

Then re-embed `dist/bundle.js` into a `<script>` tag inside a copy of
the HTML shell (see the `<style>`/`<div id="root">` wrapper already
used in the delivered `Sidebar.html`) and paste the result back into
the Apps Script editor's `Sidebar.html` file.

## Still outside this build (known gaps, being upfront)

- No two-step approval enforcement lives on the Sheets side yet - it's
  enforced Excel-side (`approvalStage !== "approved"` is rejected in
  `doPost`), but a Finance Manager can't currently approve anything
  *from* the sidebar itself. That would be a natural next feature.
- No WebSocket/push layer - Excel still polls this Web App every 6s.
- This has not been tested against a live deployment yet. The first
  real run may surface small issues (e.g. Apps Script's exact quota
  limits, sheet permission edge cases) that only show up in production.
  