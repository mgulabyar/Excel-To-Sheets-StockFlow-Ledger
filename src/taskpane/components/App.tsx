// declare const Excel: any;
// import * as React from "react";
// import { useState, useEffect } from "react";
// import {
//   Box,
//   Button,
//   TextField,
//   Typography,
//   Alert,
//   Snackbar,
// } from "@mui/material";
// import SyncAltIcon from "@mui/icons-material/SyncAlt";
// import CloudUploadIcon from "@mui/icons-material/CloudUpload";

// export default function App() {
//   const [itemCode, setItemCode] = useState<string>("");
//   const [quantity, setQuantity] = useState<string>("");

//   const [loading, setLoading] = useState<boolean>(false);
//   const [snackbarMsg, setSnackbarMsg] = useState<string>("");
//   const [openSnackbar, setOpenSnackbar] = useState<boolean>(false);

//   const GOOGLE_SCRIPT_API_URL = "YOUR_DEPLOYED_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE";

//   useEffect(() => {
//     const trackingPollingInterval = setInterval(() => {
//       fetchLatestCloudInventory();
//     }, 6000);

//     return () => clearInterval(trackingPollingInterval);
//   }, []);

//   const fetchLatestCloudInventory = async () => {
//     try {
//       const response = await fetch(`${GOOGLE_SCRIPT_API_URL}?action=readLatest`);
//       if (!response.ok) return;

//       const payloadData = await response.json();

//       if (payloadData && payloadData.lowStockItems) {
//         await Excel.run(
//           async (context: {
//             workbook: { worksheets: { getActiveWorksheet: () => any } };
//             sync: () => any;
//           }) => {
//             const sheet = context.workbook.worksheets.getActiveWorksheet();
//             const usedRange = sheet.getUsedRange();
//             usedRange.load("values");
//             await context.sync();

//             const dataValues = usedRange.values;

//             payloadData.lowStockItems.forEach((lowItemCode: string) => {
//               for (let i = 1; i < dataValues.length; i++) {
//                 if (String(dataValues[i][0]).trim() === lowItemCode.trim()) {
//                   const targetRowRange = sheet.getRow(i);
//                   const visualDataGrid = targetRowRange.getResizedRange(0, 4);
//                   visualDataGrid.format.fill.color = "#FEE2E2";
//                 }
//               }
//             });
//             await context.sync();
//           }
//         );
//       }
//     } catch (err) {
//       console.error("Automated tracking background refresh sequence failed:", err);
//     }
//   };

//   const handlePushWritebackMatrix = async () => {
//     if (!itemCode || !quantity) {
//       setSnackbarMsg("Operational Error: Both payload reference parameters are mandatory.");
//       setOpenSnackbar(true);
//       return;
//     }

//     setLoading(true);
//     try {
//       const transmissionPayload = {
//         itemCode: itemCode.trim(),
//         newQuantity: parseInt(quantity, 10),
//         timestamp: new Date().toISOString(),
//         triggeredBy: "Excel Head Office System",
//       };

//       await fetch(GOOGLE_SCRIPT_API_URL, {
//         method: "POST",
//         mode: "no-cors",
//         headers: { "Content-Type": "application/json" },
//         body: JSON.stringify(transmissionPayload),
//       });

//       await Excel.run(
//         async (context: {
//           workbook: { worksheets: { getActiveWorksheet: () => any } };
//           sync: () => any;
//         }) => {
//           const sheet = context.workbook.worksheets.getActiveWorksheet();
//           const usedRange = sheet.getUsedRange();
//           usedRange.load("values");
//           await context.sync();

//           const dataValues = usedRange.values;
//           for (let i = 1; i < dataValues.length; i++) {
//             if (String(dataValues[i][0]).trim() === itemCode.trim()) {
//               sheet.getRangeByIndexes(i, 2).values = [[parseInt(quantity, 10)]];
//               sheet.getRangeByIndexes(i, 4).values = [["Synced to Cloud Ledger"]];

//               const targetRowRange = sheet.getRow(i);
//               const visualDataGrid = targetRowRange.getResizedRange(0, 4);
//               visualDataGrid.format.fill.clear();
//               break;
//             }
//           }
//           await context.sync();
//         }
//       );

//       setSnackbarMsg("Success: Procurement data vector synchronized to StockFlow cloud ledger!");
//       setItemCode("");
//       setQuantity("");
//     } catch (err) {
//       setSnackbarMsg("Transmission Fault: Secure ledger mapping sequence failed.");
//     } finally {
//       setLoading(false);
//       setOpenSnackbar(true);
//     }
//   };

//   return (
//     <Box
//       sx={{
//         minHeight: "100vh",
//         bgcolor: "#fbfcfe",
//         p: 3,
//         display: "flex",
//         flexDirection: "column",
//         gap: 3,
//         fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif"
//       }}
//     >
//       {/* Premium Minimalist Typography Branding */}
//       <Box sx={{ mt: 1 }}>
//         <Typography
//           variant="h5"
//           sx={{
//             fontWeight: 700,
//             color: "#0062D6",
//             letterSpacing: "-0.02em",
//             fontSize: "24px"
//           }}
//         >
//           StockFlow Ledger
//         </Typography>
//         <Typography
//           variant="body2"
//           sx={{
//             color: "#64748b",
//             fontSize: "12px",
//             mt: 0.5,
//             lineHeight: 1.5
//           }}
//         >
//           Real-time enterprise middleware bridging local desktop spreadsheet architectures directly with cloud warehousing ledgers.
//         </Typography>
//       </Box>

//       {/* Modern Status Monitor */}
//       <Alert
//         severity="info"
//         icon={<SyncAltIcon fontSize="small" sx={{ color: "#0062D6" }} />}
//         sx={{
//           borderRadius: "10px",
//           fontSize: "11px",
//           py: 0.2,
//           bgcolor: "#eff6ff",
//           color: "#1e40af",
//           border: "1px solid #bfdbfe",
//           fontWeight: 500
//         }}
//       >
//         Telemetry Core Active: Checking cloud state logs (6s loop)
//       </Alert>

//       {/* Flat Content Layer (All box lines and cards removed) */}
//       <Box sx={{ display: "flex", flexDirection: "column", gap: 3.5, mt: 1 }}>
//         <Box>
//           <Typography
//             variant="subtitle2"
//             sx={{
//               fontWeight: 600,
//               color: "#1e293b",
//               fontSize: "14px",
//               mb: 0.5
//             }}
//           >
//             Procurement Overrides
//           </Typography>
//           <Typography
//             variant="caption"
//             sx={{
//               color: "#94a3b8",
//               display: "block",
//               mb: 2,
//               fontSize: "11px"
//             }}
//           >
//             Authorize instant writeback sequences to balance multi-tenant supply limits.
//           </Typography>
//         </Box>

//         <Box sx={{ display: "flex", flexDirection: "column", gap: 2.5 }}>
//           <TextField
//             label="Target Item Anchor Key"
//             variant="outlined"
//             size="small"
//             placeholder="e.g., ITEM-001"
//             value={itemCode}
//             onChange={(e) => setItemCode(e.target.value)}
//             fullWidth
//             sx={{
//               "& .MuiOutlinedInput-root": {
//                 borderRadius: "10px",
//                 bgcolor: "#ffffff"
//               }
//             }}
//             slotProps={{
//               inputLabel: { style: { fontSize: 13, color: "#64748b" } },
//               input: { style: { fontSize: 13, fontFamily: "monospace", color: "#1e293b" } },
//             }}
//           />
          
//           <TextField
//             label="Restock Vector Quantity"
//             type="number"
//             variant="outlined"
//             size="small"
//             placeholder="0"
//             value={quantity}
//             onChange={(e) => setQuantity(e.target.value)}
//             fullWidth
//             sx={{
//               "& .MuiOutlinedInput-root": {
//                 borderRadius: "10px",
//                 bgcolor: "#ffffff"
//               }
//             }}
//             slotProps={{
//               inputLabel: { style: { fontSize: 13, color: "#64748b" } },
//               input: { style: { fontSize: 13, color: "#1e293b" } },
//             }}
//           />

//           <Button
//             variant="contained"
//             disableElevation
//             disabled={loading}
//             onClick={handlePushWritebackMatrix}
//             startIcon={<CloudUploadIcon />}
//             sx={{
//               bgcolor: "#0062D6",
//               "&:hover": { bgcolor: "#0B3C95" },
//               borderRadius: "10px",
//               textTransform: "none",
//               fontWeight: 600,
//               fontSize: "13px",
//               py: 1.2,
//               boxShadow: "0 4px 12px rgba(0, 98, 214, 0.15)",
//               mt: 1
//             }}
//           >
//             {loading ? "Processing Encryption Sync..." : "Push Approval to Cloud"}
//           </Button>
//           <Typography
//             variant="caption"
//             sx={{
//               color: "#94a3b8",
//               textAlign: "center",
//               fontSize: "10px",
//               display: "block",
//               mt: 0.5
//             }}
//           >
//             Secure pipeline connection powered by StockFlow Ledger v1.0
//           </Typography>
//         </Box>
//       </Box>

//       <Snackbar
//         open={openSnackbar}
//         autoHideDuration={4000}
//         onClose={() => setOpenSnackbar(false)}
//         message={snackbarMsg}
//         anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
//       />
//     </Box>
//   );
// }


// ============================================================================
// App.tsx - StockFlow Ledger (Excel Side) - v2.0 Enterprise Edition
//
// WHAT CHANGED FROM v1:
// - Real error handling (no more silent "no-cors" fake success)
// - HMAC-signed requests (secureRequest.ts)
// - Offline retry queue (syncQueue.ts)
// - Conflict detection between local edits and cloud edits (conflictResolver.ts)
// - Full audit trail, local + written into a hidden workbook sheet (auditLogger.ts)
// - Two-step approval chain: Director submits -> Finance approves (approvalWorkflow.ts)
// - Trend chart + adaptive vendor lead-time learning (analyticsEngine.ts)
// - All Office.js calls extracted into excelController.ts (clean architecture)
// ============================================================================

import * as React from "react";
import { useState, useEffect, useCallback, useRef } from "react";
import { Box, Button, TextField, Typography, Alert, Snackbar } from "@mui/material";
import SyncAltIcon from "@mui/icons-material/SyncAlt";
import CloudUploadIcon from "@mui/icons-material/CloudUpload";

import { ApprovalStage, SyncStatus, CloudInventoryPayload } from "../types";
import { sendSecureGet } from "../utils/secureRequest";
import {
  bindSelectionHandler,
  applyStockoutHighlight,
  applyConflictHighlight,
  writeRowUpdate,
  appendAuditRowToWorkbook,
} from "../utils/excelController";
import { enqueueFailedRequest, flushQueue, getQueueLength } from "../utils/syncQueue";
import { detectConflict, resolveConflict } from "../utils/conflictResolver";
import { recordAuditEntry, getAuditLog, exportAuditLogAsCsv } from "../utils/auditLogger";
import {
  submitForFinanceApproval,
  financeApprove,
  financeReject,
  isReadyToSync,
} from "../utils/approvalWorkflow";
import { recordTrendPoint, getTrendPoints, getAdaptiveLeadTime } from "../utils/analyticsEngine";

import StatusIndicator from "../components/StatusIndicator";
import ApprovalPanel from "../components/ApprovalPanel";
import TrendChart from "../components/TrendChart";
import AuditLogPanel from "../components/AuditLogPanel";

const GOOGLE_SCRIPT_API_URL = "YOUR_DEPLOYED_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE";
const POLL_INTERVAL_MS = 6000;
const CURRENT_USER = "Excel Head Office System"; // replace with real Office.context user later

export default function App() {
  const [itemCode, setItemCode] = useState<string>("");
  const [quantity, setQuantity] = useState<string>("");

  const [loading, setLoading] = useState<boolean>(false);
  const [snackbarMsg, setSnackbarMsg] = useState<string>("");
  const [openSnackbar, setOpenSnackbar] = useState<boolean>(false);

  const [syncStatus, setSyncStatus] = useState<SyncStatus>("idle");
  const [queueLength, setQueueLength] = useState<number>(0);
  const [approvalStage, setApprovalStage] = useState<ApprovalStage>("draft");
  const [auditEntries, setAuditEntries] = useState(getAuditLog(20));

  const lastSyncedTimestampRef = useRef<string | null>(null);
  const localEditTimestampRef = useRef<string>(new Date().toISOString());

  const refreshQueueBadge = useCallback(() => {
    setQueueLength(getQueueLength());
  }, []);

  const refreshAuditPanel = useCallback(() => {
    setAuditEntries(getAuditLog(20));
  }, []);

  // --------------------------------------------------------------------
  // Setup: selection listener + polling loop + initial queue flush
  // --------------------------------------------------------------------
  useEffect(() => {
    let disposeSelectionHandler: (() => void) | null = null;

    bindSelectionHandler((code, qty) => {
      setItemCode(code);
      setQuantity(qty);
      localEditTimestampRef.current = new Date().toISOString();
    })
      .then((dispose) => {
        disposeSelectionHandler = dispose;
      })
      .catch(() => {
        // If binding fails, manual entry still works - not a hard failure
      });

    const pollingInterval = setInterval(() => {
      fetchLatestCloudInventory();
      attemptQueueFlush();
    }, POLL_INTERVAL_MS);

    refreshQueueBadge();
    refreshAuditPanel();

    return () => {
      clearInterval(pollingInterval);
      if (disposeSelectionHandler) disposeSelectionHandler();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // --------------------------------------------------------------------
  // Retry queue flush - runs every poll cycle if items are pending
  // --------------------------------------------------------------------
  const attemptQueueFlush = useCallback(async () => {
    if (getQueueLength() === 0) return;
    setSyncStatus("syncing");
    const result = await flushQueue(GOOGLE_SCRIPT_API_URL);
    refreshQueueBadge();

    if (result.succeeded > 0) {
      setSnackbarMsg(`Reconnected: ${result.succeeded} queued update(s) synced successfully.`);
      setOpenSnackbar(true);
    }
    setSyncStatus(getQueueLength() > 0 ? "offline" : "synced");
  }, [refreshQueueBadge]);

  // --------------------------------------------------------------------
  // PULL ENGINE: fetch cloud data, run predictive formula, detect conflicts
  // --------------------------------------------------------------------
  const fetchLatestCloudInventory = useCallback(async () => {
    setSyncStatus((prev) => (prev === "offline" ? prev : "syncing"));

    const result = await sendSecureGet<{ inventoryData: CloudInventoryPayload[] }>(
      `${GOOGLE_SCRIPT_API_URL}?action=readLatest`
    );

    if (!result.ok || !result.data) {
      setSyncStatus("offline");
      return;
    }

    const inventoryData = result.data.inventoryData || [];

    for (const cloudItem of inventoryData) {
      // Adaptive lead time replaces the static promised lead time
      const adaptiveLeadTime = getAdaptiveLeadTime(cloudItem.itemCode, cloudItem.leadTime);

      const totalProjectedStock = cloudItem.physicalStock + cloudItem.expectedCargo;
      const absoluteRiskThreshold = adaptiveLeadTime * cloudItem.dailySalesRate;
      const isStockoutRisk = totalProjectedStock <= absoluteRiskThreshold;

      // Conflict check: was this item also edited locally since last sync?
      if (cloudItem.itemCode === itemCode && cloudItem.lastUpdated) {
        const conflict = detectConflict({
          itemCode: cloudItem.itemCode,
          localValue: Number(quantity) || 0,
          localTimestamp: localEditTimestampRef.current,
          cloudValue: cloudItem.physicalStock,
          cloudTimestamp: cloudItem.lastUpdated,
          lastSyncedTimestamp: lastSyncedTimestampRef.current,
        });

        if (conflict) {
          const resolved = resolveConflict(conflict, "manual_for_flagged", isStockoutRisk);
          await applyConflictHighlight(cloudItem.itemCode);
          setSyncStatus("conflict");
          recordAuditEntry({
            itemCode: cloudItem.itemCode,
            action: "conflict_resolved",
            beforeValue: conflict.localValue,
            afterValue: conflict.cloudValue,
            user: CURRENT_USER,
          });
          refreshAuditPanel();
          continue; // don't apply stockout highlight over an unresolved conflict
        }
      }

      await applyStockoutHighlight(cloudItem.itemCode, isStockoutRisk);
      recordTrendPoint(cloudItem.itemCode, cloudItem.physicalStock);
    }

    lastSyncedTimestampRef.current = new Date().toISOString();
    setSyncStatus(getQueueLength() > 0 ? "offline" : "synced");
  }, [itemCode, quantity, refreshAuditPanel]);

  // --------------------------------------------------------------------
  // Approval chain handlers
  // --------------------------------------------------------------------
  const handleSubmitForApproval = () => {
    const result = submitForFinanceApproval(approvalStage);
    if (result.allowed) {
      setApprovalStage(result.nextStage);
      recordAuditEntry({
        itemCode: itemCode || "N/A",
        action: "approval_stage_change",
        beforeValue: "draft",
        afterValue: result.nextStage,
        user: CURRENT_USER,
      });
      refreshAuditPanel();
    }
    setSnackbarMsg(result.message);
    setOpenSnackbar(true);
  };

  const handleFinanceApprove = () => {
    const result = financeApprove(approvalStage);
    if (result.allowed) {
      setApprovalStage(result.nextStage);
      recordAuditEntry({
        itemCode: itemCode || "N/A",
        action: "approval_stage_change",
        beforeValue: "pending_finance",
        afterValue: result.nextStage,
        user: "Finance Manager",
      });
      refreshAuditPanel();
    }
    setSnackbarMsg(result.message);
    setOpenSnackbar(true);
  };

  const handleFinanceReject = () => {
    const result = financeReject(approvalStage);
    if (result.allowed) {
      setApprovalStage(result.nextStage);
      recordAuditEntry({
        itemCode: itemCode || "N/A",
        action: "approval_stage_change",
        beforeValue: "pending_finance",
        afterValue: result.nextStage,
        user: "Finance Manager",
      });
      refreshAuditPanel();
    }
    setSnackbarMsg(result.message);
    setOpenSnackbar(true);
  };

  // --------------------------------------------------------------------
  // PUSH ENGINE: secure, queued, audited writeback
  // --------------------------------------------------------------------
  const handlePushWritebackMatrix = async () => {
    if (!itemCode || !quantity) {
      setSnackbarMsg("Operational Error: Both payload reference parameters are mandatory.");
      setOpenSnackbar(true);
      return;
    }

    if (!isReadyToSync(approvalStage)) {
      setSnackbarMsg("Blocked: this writeback needs Finance approval before it can sync.");
      setOpenSnackbar(true);
      return;
    }

    setLoading(true);
    setSyncStatus("syncing");

    const transmissionPayload = {
      itemCode: itemCode.trim(),
      newQuantity: parseInt(quantity, 10),
      timestamp: new Date().toISOString(),
      triggeredBy: CURRENT_USER,
      approvalStage: approvalStage,
    };

    try {
      const { sendSecurePost } = await import("../utils/secureRequest");
      const result = await sendSecurePost(GOOGLE_SCRIPT_API_URL, transmissionPayload);

      if (!result.ok) {
        // Real failure now correctly detected (fixes the old no-cors bug)
        enqueueFailedRequest(transmissionPayload);
        refreshQueueBadge();
        setSyncStatus("offline");
        setSnackbarMsg(`Sync failed - queued for retry. (${result.errorMessage})`);
        setOpenSnackbar(true);
        return;
      }

      const { previousQuantity } = await writeRowUpdate(
        itemCode,
        parseInt(quantity, 10),
        "Synced to Cloud Ledger"
      );

      recordAuditEntry({
        itemCode: itemCode.trim(),
        action: "push_approval",
        beforeValue: previousQuantity,
        afterValue: parseInt(quantity, 10),
        user: CURRENT_USER,
      });

      await appendAuditRowToWorkbook([
        `audit-${Date.now()}`,
        itemCode.trim(),
        "push_approval",
        previousQuantity ?? "-",
        parseInt(quantity, 10),
        new Date().toISOString(),
      ]);

      refreshAuditPanel();
      setSyncStatus("synced");
      setSnackbarMsg("Success: Procurement data vector synchronized to StockFlow cloud ledger!");
      setItemCode("");
      setQuantity("");
      setApprovalStage("draft");
    } catch (err) {
      enqueueFailedRequest(transmissionPayload);
      refreshQueueBadge();
      setSyncStatus("offline");
      setSnackbarMsg("Network fault: request queued locally and will retry automatically.");
    } finally {
      setLoading(false);
      setOpenSnackbar(true);
    }
  };

  const trendPoints = itemCode ? getTrendPoints(itemCode) : [];

  return (
    <Box
      sx={{
        minHeight: "100vh",
        bgcolor: "#fbfcfe",
        p: 3,
        display: "flex",
        flexDirection: "column",
        gap: 3,
        fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
      }}
    >
      {/* Branding */}
      <Box sx={{ mt: 1 }}>
        <Typography variant="h5" sx={{ fontWeight: 700, color: "#0062D6", letterSpacing: "-0.02em", fontSize: "24px" }}>
          StockFlow Ledger
        </Typography>
        <Typography variant="body2" sx={{ color: "#64748b", fontSize: "12px", mt: 0.5, lineHeight: 1.5 }}>
          Near-instant enterprise middleware bridging local desktop spreadsheet architectures with cloud warehousing ledgers.
        </Typography>
      </Box>

      {/* Live status row (replaces the static "Telemetry Core Active" alert) */}
      <Box sx={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <StatusIndicator status={syncStatus} queueLength={queueLength} />
      </Box>

      <Alert
        severity="info"
        icon={<SyncAltIcon fontSize="small" sx={{ color: "#0062D6" }} />}
        sx={{
          borderRadius: "10px",
          fontSize: "11px",
          py: 0.2,
          bgcolor: "#eff6ff",
          color: "#1e40af",
          border: "1px solid #bfdbfe",
          fontWeight: 500,
        }}
      >
        Sync loop active - checking cloud state every {POLL_INTERVAL_MS / 1000}s (near-instant, not push-based yet)
      </Alert>

      {/* Procurement input + approval chain */}
      <Box sx={{ display: "flex", flexDirection: "column", gap: 2.5, mt: 1 }}>
        <Box>
          <Typography variant="subtitle2" sx={{ fontWeight: 600, color: "#1e293b", fontSize: "14px", mb: 0.5 }}>
            Procurement Overrides
          </Typography>
          <Typography variant="caption" sx={{ color: "#94a3b8", display: "block", mb: 2, fontSize: "11px" }}>
            Authorize instant writeback sequences to balance multi-tenant supply limits.
          </Typography>
        </Box>

        <TextField
          label="Target Item Anchor Key"
          variant="outlined"
          size="small"
          placeholder="e.g., ITEM-001"
          value={itemCode}
          onChange={(e) => {
            setItemCode(e.target.value);
            localEditTimestampRef.current = new Date().toISOString();
          }}
          fullWidth
          sx={{ "& .MuiOutlinedInput-root": { borderRadius: "10px", bgcolor: "#ffffff" } }}
        />

        <TextField
          label="Restock Vector Quantity"
          type="number"
          variant="outlined"
          size="small"
          placeholder="0"
          value={quantity}
          onChange={(e) => {
            setQuantity(e.target.value);
            localEditTimestampRef.current = new Date().toISOString();
          }}
          fullWidth
          sx={{ "& .MuiOutlinedInput-root": { borderRadius: "10px", bgcolor: "#ffffff" } }}
        />

        <ApprovalPanel
          stage={approvalStage}
          onSubmitForApproval={handleSubmitForApproval}
          onFinanceApprove={handleFinanceApprove}
          onFinanceReject={handleFinanceReject}
          disabled={loading}
        />

        <Button
          variant="contained"
          disableElevation
          disabled={loading || !isReadyToSync(approvalStage)}
          onClick={handlePushWritebackMatrix}
          startIcon={<CloudUploadIcon />}
          sx={{
            bgcolor: "#0062D6",
            "&:hover": { bgcolor: "#0B3C95" },
            borderRadius: "10px",
            textTransform: "none",
            fontWeight: 600,
            fontSize: "13px",
            py: 1.2,
            boxShadow: "0 4px 12px rgba(0, 98, 214, 0.15)",
            mt: 1,
          }}
        >
          {loading ? "Processing Encryption Sync..." : "Push Approval to Cloud"}
        </Button>

        <Typography variant="caption" sx={{ color: "#94a3b8", textAlign: "center", fontSize: "10px", display: "block" }}>
          Secure HMAC-signed pipeline · StockFlow Ledger v2.0
        </Typography>

        <TrendChart itemCode={itemCode} points={trendPoints} />

        <AuditLogPanel
          entries={auditEntries}
          onExportCsv={() => {
            const csv = exportAuditLogAsCsv();
            const blob = new Blob([csv], { type: "text/csv" });
            const url = URL.createObjectURL(blob);
            const link = document.createElement("a");
            link.href = url;
            link.download = "stockflow_audit_log.csv";
            link.click();
            URL.revokeObjectURL(url);
          }}
        />
      </Box>

      <Snackbar
        open={openSnackbar}
        autoHideDuration={4000}
        onClose={() => setOpenSnackbar(false)}
        message={snackbarMsg}
        anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
      />
    </Box>
  );
}
