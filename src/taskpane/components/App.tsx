declare const Excel: any;
import * as React from "react";
import { Box, Typography, Button, TextField, CircularProgress } from "@mui/material";
import SyncAltIcon from "@mui/icons-material/SyncAlt";

const GOOGLE_SCRIPT_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzE7ehrPGVe3eKK6qGI8ZdgAOjy5u_rftI_av8wdXCfb-mzWhs73EvUCKvizM5lXFcoQg/exec";

export default function App() {
  const [itemCode, setItemCode] = React.useState("");
  const [quantity, setQuantity] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [toast, setToast] = React.useState<{ type: "success" | "error"; msg: string } | null>(null);

  const showToast = (type: "success" | "error", msg: string) => {
    setToast({ type, msg });
    setTimeout(() => setToast(null), 3000);
  };

  const pushToSheets = async () => {
    if (!itemCode || !quantity) {
      showToast("error", "Fill both fields!");
      return;
    }

    setLoading(true);
    setToast(null);

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
        showToast("success", "Data sent!");
        setItemCode("");
        setQuantity("");
      } else {
        throw new Error(result.message || "Unknown error");
      }
    } catch (error: any) {
      showToast("error", error.message);
    } finally {
      setLoading(false);
    }
  };

  const fetchFromSheets = async () => {
    setLoading(true);
    setToast(null);

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
            throw new Error("Invalid data structure");
          }
          
          const range = sheet.getRangeByIndexes(0, 0, rowCount, colCount);
          range.values = result.data;
          range.format.autofitColumns();
          await context.sync();
        });

        showToast("success", "Data loaded!");
      } else {
        showToast("error", "No data found");
      }
    } catch (error: any) {
      showToast("error", error.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <Box sx={{ 
      p: 2, 
      display: "flex", 
      flexDirection: "column", 
      gap: 2, 
      bgcolor: "#fbfcfe", 
      minHeight: "100vh",
    }}>
      {/* Floating Toast */}
      {toast && (
        <Box
          sx={{
            position: "fixed",
            top: 12,
            right: 12,
            zIndex: 9999,
            minWidth: 160,
            animation: "slideIn 0.3s ease-out",
            animationFillMode: "forwards",
          }}
        >
          <Box
            sx={{
              px: 2,
              py: 0.75,
              borderRadius: 1,
              fontSize: "11px",
              fontWeight: 600,
              color: toast.type === "success" ? "#155724" : "#721c24",
              bgcolor: toast.type === "success" ? "#d4edda" : "#f8d7da",
              border: `1px solid ${toast.type === "success" ? "#c3e6cb" : "#f5c6cb"}`,
              boxShadow: "0 6px 16px rgba(0,0,0,0.12)",
              animation: "slideOut 0.3s ease-in 2.7s",
              animationFillMode: "forwards",
            }}
          >
            {toast.msg}
          </Box>
        </Box>
      )}

      {/* Header */}
      <Box sx={{ textAlign: "center", mb: 1 }}>
        <Typography sx={{ fontWeight: 700, color: "#0062D6", fontSize: "19px", letterSpacing: "0.3px" }}>
          StockFlow Ledger
        </Typography>
        <Typography 
          sx={{ 
            color: "#64748b", 
            display: "flex", 
            justifyContent: "center", 
            alignItems: "center", 
            gap: 1,
            fontSize: "12px",
            fontWeight: 500,
          }}
        >
          Excel Add-ins <SyncAltIcon sx={{ fontSize: "15px" }} /> Google Sheets
        </Typography>
      </Box>

      {/* Inputs */}
      <TextField
        label="Item Name"
        variant="outlined"
        size="small"
        value={itemCode}
        onChange={(e) => setItemCode(e.target.value)}
        fullWidth
        sx={{
          "& .MuiOutlinedInput-root": {
            fontSize: "13px",
            "& fieldset": {
              borderColor: "#e2e8f0",
            },
            "&:hover fieldset": {
              borderColor: "#0062D6",
            },
            "&.Mui-focused fieldset": {
              borderColor: "#0062D6",
            },
          },
          "& .MuiInputLabel-root": {
            fontSize: "13px",
            color: "#64748b",
          },
        }}
      />

      <TextField
        label="Quantity"
        type="number"
        variant="outlined"
        size="small"
        value={quantity}
        onChange={(e) => setQuantity(e.target.value)}
        fullWidth
        sx={{
          "& .MuiOutlinedInput-root": {
            fontSize: "13px",
            "& fieldset": {
              borderColor: "#e2e8f0",
            },
            "&:hover fieldset": {
              borderColor: "#0062D6",
            },
            "&.Mui-focused fieldset": {
              borderColor: "#0062D6",
            },
          },
          "& .MuiInputLabel-root": {
            fontSize: "13px",
            color: "#64748b",
          },
        }}
      />

      {/* Buttons */}
      <Button
        variant="contained"
        disabled={loading}
        onClick={pushToSheets}
        fullWidth
        sx={{ 
          bgcolor: "#0062D6", 
          textTransform: "none", 
          py: 0.75,
          fontSize: "13px",
          fontWeight: 600,
          letterSpacing: "0.3px",
          boxShadow: "0 2px 8px rgba(0,98,214,0.25)",
          "&:hover": { 
            bgcolor: "#0B3C95",
            boxShadow: "0 4px 12px rgba(0,98,214,0.35)",
          },
          "&.Mui-disabled": {
            bgcolor: "#94a3b8",
            color: "#e2e8f0",
          },
        }}
      >
        {loading ? <CircularProgress size={18} color="inherit" sx={{ mr: 1 }} /> : "Send to Sheets"}
      </Button>

      <Button
        variant="outlined"
        disabled={loading}
        onClick={fetchFromSheets}
        fullWidth
        sx={{ 
          textTransform: "none", 
          py: 0.75,
          fontSize: "13px",
          fontWeight: 600,
          letterSpacing: "0.3px",
          color: "#0062D6",
          borderColor: "#0062D6",
          "&:hover": { 
            borderColor: "#0B3C95", 
            bgcolor: "rgba(0,98,214,0.04)",
          },
          "&.Mui-disabled": {
            color: "#94a3b8",
            borderColor: "#94a3b8",
          },
        }}
      >
        Fetch from Sheets
      </Button>
    </Box>
  );
}