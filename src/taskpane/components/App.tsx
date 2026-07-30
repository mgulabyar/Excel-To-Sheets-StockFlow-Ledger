declare const Excel: any;
import * as React from "react";
import { useState, useEffect } from "react";
import {
  Box,
  Button,
  TextField,
  Typography,
  Alert,
  Snackbar,
} from "@mui/material";
import SyncAltIcon from "@mui/icons-material/SyncAlt";
import CloudUploadIcon from "@mui/icons-material/CloudUpload";

export default function App() {
  const [itemCode, setItemCode] = useState<string>("");
  const [quantity, setQuantity] = useState<string>("");

  const [loading, setLoading] = useState<boolean>(false);
  const [snackbarMsg, setSnackbarMsg] = useState<string>("");
  const [openSnackbar, setOpenSnackbar] = useState<boolean>(false);

  const GOOGLE_SCRIPT_API_URL = "YOUR_DEPLOYED_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE";

  useEffect(() => {
    const trackingPollingInterval = setInterval(() => {
      fetchLatestCloudInventory();
    }, 6000);

    return () => clearInterval(trackingPollingInterval);
  }, []);

  const fetchLatestCloudInventory = async () => {
    try {
      const response = await fetch(`${GOOGLE_SCRIPT_API_URL}?action=readLatest`);
      if (!response.ok) return;

      const payloadData = await response.json();

      if (payloadData && payloadData.lowStockItems) {
        await Excel.run(
          async (context: {
            workbook: { worksheets: { getActiveWorksheet: () => any } };
            sync: () => any;
          }) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();
            usedRange.load("values");
            await context.sync();

            const dataValues = usedRange.values;

            payloadData.lowStockItems.forEach((lowItemCode: string) => {
              for (let i = 1; i < dataValues.length; i++) {
                if (String(dataValues[i][0]).trim() === lowItemCode.trim()) {
                  const targetRowRange = sheet.getRow(i);
                  const visualDataGrid = targetRowRange.getResizedRange(0, 4);
                  visualDataGrid.format.fill.color = "#FEE2E2";
                }
              }
            });
            await context.sync();
          }
        );
      }
    } catch (err) {
      console.error("Automated tracking background refresh sequence failed:", err);
    }
  };

  const handlePushWritebackMatrix = async () => {
    if (!itemCode || !quantity) {
      setSnackbarMsg("Operational Error: Both payload reference parameters are mandatory.");
      setOpenSnackbar(true);
      return;
    }

    setLoading(true);
    try {
      const transmissionPayload = {
        itemCode: itemCode.trim(),
        newQuantity: parseInt(quantity, 10),
        timestamp: new Date().toISOString(),
        triggeredBy: "Excel Head Office System",
      };

      await fetch(GOOGLE_SCRIPT_API_URL, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(transmissionPayload),
      });

      await Excel.run(
        async (context: {
          workbook: { worksheets: { getActiveWorksheet: () => any } };
          sync: () => any;
        }) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const usedRange = sheet.getUsedRange();
          usedRange.load("values");
          await context.sync();

          const dataValues = usedRange.values;
          for (let i = 1; i < dataValues.length; i++) {
            if (String(dataValues[i][0]).trim() === itemCode.trim()) {
              sheet.getRangeByIndexes(i, 2).values = [[parseInt(quantity, 10)]];
              sheet.getRangeByIndexes(i, 4).values = [["Synced to Cloud Ledger"]];

              const targetRowRange = sheet.getRow(i);
              const visualDataGrid = targetRowRange.getResizedRange(0, 4);
              visualDataGrid.format.fill.clear();
              break;
            }
          }
          await context.sync();
        }
      );

      setSnackbarMsg("Success: Procurement data vector synchronized to StockFlow cloud ledger!");
      setItemCode("");
      setQuantity("");
    } catch (err) {
      setSnackbarMsg("Transmission Fault: Secure ledger mapping sequence failed.");
    } finally {
      setLoading(false);
      setOpenSnackbar(true);
    }
  };

  return (
    <Box
      sx={{
        minHeight: "100vh",
        bgcolor: "#fbfcfe",
        p: 3,
        display: "flex",
        flexDirection: "column",
        gap: 3,
        fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif"
      }}
    >
      {/* Premium Minimalist Typography Branding */}
      <Box sx={{ mt: 1 }}>
        <Typography
          variant="h5"
          sx={{
            fontWeight: 700,
            color: "#0062D6",
            letterSpacing: "-0.02em",
            fontSize: "24px"
          }}
        >
          StockFlow Ledger
        </Typography>
        <Typography
          variant="body2"
          sx={{
            color: "#64748b",
            fontSize: "12px",
            mt: 0.5,
            lineHeight: 1.5
          }}
        >
          Real-time enterprise middleware bridging local desktop spreadsheet architectures directly with cloud warehousing ledgers.
        </Typography>
      </Box>

      {/* Modern Status Monitor */}
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
          fontWeight: 500
        }}
      >
        Telemetry Core Active: Checking cloud state logs (6s loop)
      </Alert>

      {/* Flat Content Layer (All box lines and cards removed) */}
      <Box sx={{ display: "flex", flexDirection: "column", gap: 3.5, mt: 1 }}>
        <Box>
          <Typography
            variant="subtitle2"
            sx={{
              fontWeight: 600,
              color: "#1e293b",
              fontSize: "14px",
              mb: 0.5
            }}
          >
            Procurement Overrides
          </Typography>
          <Typography
            variant="caption"
            sx={{
              color: "#94a3b8",
              display: "block",
              mb: 2,
              fontSize: "11px"
            }}
          >
            Authorize instant writeback sequences to balance multi-tenant supply limits.
          </Typography>
        </Box>

        <Box sx={{ display: "flex", flexDirection: "column", gap: 2.5 }}>
          <TextField
            label="Target Item Anchor Key"
            variant="outlined"
            size="small"
            placeholder="e.g., ITEM-001"
            value={itemCode}
            onChange={(e) => setItemCode(e.target.value)}
            fullWidth
            sx={{
              "& .MuiOutlinedInput-root": {
                borderRadius: "10px",
                bgcolor: "#ffffff"
              }
            }}
            slotProps={{
              inputLabel: { style: { fontSize: 13, color: "#64748b" } },
              input: { style: { fontSize: 13, fontFamily: "monospace", color: "#1e293b" } },
            }}
          />
          
          <TextField
            label="Restock Vector Quantity"
            type="number"
            variant="outlined"
            size="small"
            placeholder="0"
            value={quantity}
            onChange={(e) => setQuantity(e.target.value)}
            fullWidth
            sx={{
              "& .MuiOutlinedInput-root": {
                borderRadius: "10px",
                bgcolor: "#ffffff"
              }
            }}
            slotProps={{
              inputLabel: { style: { fontSize: 13, color: "#64748b" } },
              input: { style: { fontSize: 13, color: "#1e293b" } },
            }}
          />

          <Button
            variant="contained"
            disableElevation
            disabled={loading}
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
              mt: 1
            }}
          >
            {loading ? "Processing Encryption Sync..." : "Push Approval to Cloud"}
          </Button>
          <Typography
            variant="caption"
            sx={{
              color: "#94a3b8",
              textAlign: "center",
              fontSize: "10px",
              display: "block",
              mt: 0.5
            }}
          >
            Secure pipeline connection powered by StockFlow Ledger v1.0
          </Typography>
        </Box>
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
