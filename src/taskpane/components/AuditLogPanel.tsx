// 
// AuditLogPanel.tsx - Full Traceability 
// ============================================================================

import * as React from "react";
import { Box, Typography, Button } from "@mui/material";
import { AuditEntry } from "../types";

interface AuditLogPanelProps {
  entries: AuditEntry[];
  onExportCsv: () => void;
}

export default function AuditLogPanel({ entries, onExportCsv }: AuditLogPanelProps) {
  return (
    <Box
      sx={{
        border: "1px solid #e2e8f0",
        borderRadius: "10px",
        p: 1.5,
        bgcolor: "#ffffff",
      }}
    >
      <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 1 }}>
        <Typography variant="caption" sx={{ fontWeight: 600, color: "#334155", fontSize: "12px" }}>
          Recent Activity
        </Typography>
        <Button
          size="small"
          onClick={onExportCsv}
          sx={{ textTransform: "none", fontSize: "10px", minWidth: "auto", p: 0.5 }}
        >
          Export CSV
        </Button>
      </Box>

      {entries.length === 0 ? (
        <Typography variant="caption" sx={{ fontSize: "11px", color: "#94a3b8" }}>
          No activity recorded yet.
        </Typography>
      ) : (
        <Box sx={{ display: "flex", flexDirection: "column", gap: 0.75, maxHeight: 160, overflowY: "auto" }}>
          {entries.map((entry) => (
            <Box key={entry.id} sx={{ borderBottom: "1px solid #f1f5f9", pb: 0.5 }}>
              <Typography variant="caption" sx={{ fontSize: "10.5px", color: "#1e293b", display: "block" }}>
                <strong>{entry.itemCode}</strong> - {entry.action.replace(/_/g, " ")}
              </Typography>
              <Typography variant="caption" sx={{ fontSize: "9.5px", color: "#94a3b8" }}>
                {entry.beforeValue ?? "-"} → {entry.afterValue ?? "-"} · {new Date(entry.timestamp).toLocaleString()}
              </Typography>
            </Box>
          ))}
        </Box>
      )}
    </Box>
  );
}