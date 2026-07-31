import * as React from "react";
import { Box, Typography, Button, TextField, Alert, CircularProgress } from "@mui/material";
import SyncAltIcon from "@mui/icons-material/SyncAlt";


const GOOGLE_SCRIPT_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzE7ehrPGVe3eKK6qGI8ZdgAOjy5u_rftI_av8wdXCfb-mzWhs73EvUCKvizM5lXFcoQg/exec";


export default function App() {
  const [itemCode, setItemCode] = React.useState("");
  const [quantity, setQuantity] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [status, setStatus] = React.useState<{ type: "success" | "error"; msg: string } | null>(null);


  const pushToSheets = async () => {
    if (!itemCode || !quantity) {
      setStatus({ type: "error", msg: "Please fill both fields!" });
      return;
    }

    setLoading(true);
    setStatus(null);

    try {

      const response = await fetch(GOOGLE_SCRIPT_WEB_APP_URL, {
        method: "POST",
        body: JSON.stringify({ action: "writeData", item: itemCode, qty: quantity }),
        
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const result = await response.json();
      
      if (result.status === "success") {
        setStatus({ type: "success", msg: "Data sent to Google Sheets successfully!" });
        setItemCode("");
        setQuantity("");
      } else {
        throw new Error(result.message || "Unknown error from Sheets");
      }
    } catch (error: any) {
      setStatus({ type: "error", msg: "Failed to send data: " + error.message });
    } finally {
      setLoading(false);
    }
  };


  const fetchFromSheets = async () => {
    setLoading(true);
    setStatus(null);

    try {
      const response = await fetch(`${GOOGLE_SCRIPT_WEB_APP_URL}?action=readData`);
      
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const result = await response.json();
      
      if (result.status === "success" && result.data && result.data.length > 0) {
        await Excel.run(async (context: any) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          
          const rowCount = result.data.length;
          const colCount = result.data[0]?.length || 0;
          
          if (rowCount === 0 || colCount === 0) {
            throw new Error("Invalid data structure from Sheets");
          }
          
          const range = sheet.getRangeByIndexes(0, 0, rowCount, colCount);
          range.values = result.data;
          range.format.autofitColumns();
          await context.sync();
        });

        setStatus({ type: "success", msg: "Data loaded from Sheets into Excel!" });
      } else {
        setStatus({ type: "error", msg: "No data found in Google Sheets." });
      }
    } catch (error: any) {
      setStatus({ type: "error", msg: "Failed to fetch data: " + error.message });
    } finally {
      setLoading(false);
    }
  };


  return (
    <Box sx={{ p: 3, display: "flex", flexDirection: "column", gap: 3, bgcolor: "#fbfcfe", minHeight: "100vh" }}>
      <Box sx={{ textAlign: "center" }}>
        <Typography sx={{ fontWeight: "bold", color: "#0062D6", fontSize: "18px" }}>
          StockFlow Ledger
        </Typography>
        <Typography variant="body2" sx={{ color: "#64748b", display: "flex", justifyContent: "center", alignItems: "center", gap: 1 }}>
          Excel <SyncAltIcon sx={{ fontSize: "16px" }} /> Google Sheets
        </Typography>
      </Box>

      {/* Status Indicators */}
      {status && <Alert severity={status.type}>{status.msg}</Alert>}

      {/* Input Fields */}
      <TextField
        label="Item Name / Code"
        variant="outlined"
        size="small"
        value={itemCode}
        onChange={(e) => setItemCode(e.target.value)}
        fullWidth
      />

      <TextField
        label="Quantity"
        type="number"
        variant="outlined"
        size="small"
        value={quantity}
        onChange={(e) => setQuantity(e.target.value)}
        fullWidth
      />

      {/* Sync Actions */}
      <Button
        variant="contained"
        disabled={loading}
        onClick={pushToSheets}
        fullWidth
        sx={{ bgcolor: "#0062D6", textTransform: "none", py: 1 }}
      >
        {loading ? <CircularProgress size={24} color="inherit" /> : "Send to Google Sheets"}
      </Button>

      <Button
        variant="outlined"
        disabled={loading}
        onClick={fetchFromSheets}
        fullWidth
        sx={{ textTransform: "none", py: 1 }}
      >
        Fetch from Google Sheets
      </Button>
    </Box>
  );
}