// ============================================================================
// StatusIndicator.tsx - Live Sync Status Dot + Queue Badge
// ============================================================================

import * as React from "react";
import { Box, Typography, Tooltip } from "@mui/material";
import { SyncStatus } from "../types";

interface StatusIndicatorProps {
  status: SyncStatus;
  queueLength: number;
}

const STATUS_CONFIG: Record<SyncStatus, { color: string; label: string }> = {
  idle: { color: "#94a3b8", label: "Idle" },
  syncing: { color: "#eab308", label: "Syncing..." },
  synced: { color: "#22c55e", label: "Synced" },
  conflict: { color: "#f97316", label: "Conflict detected" },
  offline: { color: "#ef4444", label: "Offline - queued locally" },
  error: { color: "#ef4444", label: "Sync error" },
};

export default function StatusIndicator({ status, queueLength }: StatusIndicatorProps) {
  const config = STATUS_CONFIG[status];

  return (
    <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
      <Tooltip title={config.label}>
        <Box
          sx={{
            width: 10,
            height: 10,
            borderRadius: "50%",
            bgcolor: config.color,
            boxShadow: `0 0 0 3px ${config.color}22`,
            transition: "background-color 0.3s ease",
          }}
        />
      </Tooltip>
      <Typography variant="caption" sx={{ fontSize: "11px", color: "#475569", fontWeight: 500 }}>
        {config.label}
      </Typography>
      {queueLength > 0 && (
        <Box
          sx={{
            ml: 0.5,
            px: 0.8,
            py: 0.1,
            borderRadius: "6px",
            bgcolor: "#fef3c7",
            border: "1px solid #fde68a",
          }}
        >
          <Typography variant="caption" sx={{ fontSize: "10px", color: "#92400e", fontWeight: 600 }}>
            {queueLength} queued
          </Typography>
        </Box>
      )}
    </Box>
  );
}