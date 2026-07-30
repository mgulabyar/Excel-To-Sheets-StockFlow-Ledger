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