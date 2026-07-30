declare const Excel: any;
import * as React from "react";
import { useState, useEffect } from "react";
import {
  Box,
  Button,
  TextField,
  Typography,
  Card,
  CardContent,
  Alert,
  Snackbar,
} from "@mui/material";
import SyncAltIcon from "@mui/icons-material/SyncAlt";
import CloudUploadIcon from "@mui/icons-material/CloudUpload";
import StorageIcon from "@mui/icons-material/Storage";


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

      const serverResponse = await fetch(GOOGLE_SCRIPT_API_URL, {
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
        bgcolor: "#f8fafc",
        p: 2,
        display: "flex",
        flexDirection: "column",
        gap: 2,
      }}
    >

      <Alert
        severity="info"
        icon={<SyncAltIcon fontSize="small" />}
        sx={{ borderRadius: 3, fontSize: "11px", py: 0.5 }}
      >
        Telemetry Monitor Active: Syncing every 6s.
      </Alert>

      <Card
        sx={{
          borderRadius: 4,
          boxShadow: "0 4px 20px rgba(0,0,0,0.02)",
          border: "1px solid #f1f5f9",
        }}
      >
        <CardContent sx={{ p: 2.5 }}>
          <Box sx={{ display: "flex", alignItems: "center", gap: 1, mb: 2 }}>
            <StorageIcon color="primary" sx={{ fontSize: 20 }} />
            <Typography
              variant="subtitle2"
              sx={{ fontWeight: 700, color: "#1e293b", tracking: "-0.01em" }}
            >
              Procurement Overrides
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
              slotProps={{
                inputLabel: { style: { fontSize: 13 } },
                input: { style: { fontSize: 13, fontFamily: "monospace" } },
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
              slotProps={{
                inputLabel: { style: { fontSize: 13 } },
                input: { style: { fontSize: 13 } },
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
                borderRadius: 3,
                textTransform: "none",
                fontWeight: 600,
                fontSize: "13px",
                py: 1,
              }}
            >
              {loading ? "Processing Sync..." : "Push Approval to Cloud"}
            </Button>
          </Box>
        </CardContent>
      </Card>

      {/* System Status Pipeline Alert Bar Container */}
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
