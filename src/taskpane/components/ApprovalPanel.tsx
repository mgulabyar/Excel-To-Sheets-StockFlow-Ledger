
// ApprovalPanel.tsx - Two-Step Approval Chain UI
// 

import * as React from "react";
import { Box, Button, Typography, Chip } from "@mui/material";
import { ApprovalStage } from "../types";
import { stageLabel } from "../utils/approvalWorkflow";

interface ApprovalPanelProps {
  stage: ApprovalStage;
  onSubmitForApproval: () => void;
  onFinanceApprove: () => void;
  onFinanceReject: () => void;
  disabled: boolean;
}

const STAGE_COLOR: Record<ApprovalStage, "default" | "warning" | "success" | "error"> = {
  draft: "default",
  pending_finance: "warning",
  approved: "success",
  rejected: "error",
};

export default function ApprovalPanel({
  stage,
  onSubmitForApproval,
  onFinanceApprove,
  onFinanceReject,
  disabled,
}: ApprovalPanelProps) {
  return (
    <Box
      sx={{
        p: 1.5,
        borderRadius: "10px",
        bgcolor: "#f8fafc",
        border: "1px solid #e2e8f0",
        display: "flex",
        flexDirection: "column",
        gap: 1,
      }}
    >
      <Box sx={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <Typography variant="caption" sx={{ fontWeight: 600, color: "#334155", fontSize: "12px" }}>
          Approval Status
        </Typography>
        <Chip label={stageLabel(stage)} color={STAGE_COLOR[stage]} size="small" sx={{ fontSize: "10px", height: 20 }} />
      </Box>

      {stage === "draft" && (
        <Button
          size="small"
          variant="outlined"
          disabled={disabled}
          onClick={onSubmitForApproval}
          sx={{ textTransform: "none", fontSize: "12px", borderRadius: "8px" }}
        >
          Submit for Finance Approval
        </Button>
      )}

      {stage === "pending_finance" && (
        <Box sx={{ display: "flex", gap: 1 }}>
          <Button
            size="small"
            variant="contained"
            color="success"
            disabled={disabled}
            onClick={onFinanceApprove}
            sx={{ textTransform: "none", fontSize: "12px", borderRadius: "8px", flex: 1 }}
          >
            Finance Approve
          </Button>
          <Button
            size="small"
            variant="outlined"
            color="error"
            disabled={disabled}
            onClick={onFinanceReject}
            sx={{ textTransform: "none", fontSize: "12px", borderRadius: "8px", flex: 1 }}
          >
            Reject
          </Button>
        </Box>
      )}

      {stage === "approved" && (
        <Typography variant="caption" sx={{ fontSize: "11px", color: "#16a34a" }}>
          Ready - this item will sync to the cloud ledger on next push.
        </Typography>
      )}

      {stage === "rejected" && (
        <Typography variant="caption" sx={{ fontSize: "11px", color: "#dc2626" }}>
          Blocked - this writeback will not be sent to the cloud ledger.
        </Typography>
      )}
    </Box>
  );
}